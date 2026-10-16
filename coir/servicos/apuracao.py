"""Banco -> entradas do motor fiscal -> apuração (o motor não conhece o banco).

Liga cada operação ao ativo (ativo_id da F3), ao tipo de ativo (data/tipos_ativo.csv), ao
vencimento da série (prazo impresso "MM/AA" + regra de data/vencimentos_opcoes.csv) e à opção
exercida; a compra a termo recebe a data de liquidação (pregão + prazo). As linhas BM&F viram
ajustes de futuros com o valor impresso.
"""

import re
from datetime import date
from decimal import Decimal
from pathlib import Path

from sqlalchemy import select
from sqlalchemy.orm import Session, selectinload

from coir.db.modelos import Conta, EventoCorporativo, NotaCorretagem, Operacao, TickerHistorico
from coir.db.sessao import conectar
from coir.domain.fiscal import (
    AjusteFuturo,
    Apuracao,
    Evento,
    Grupo,
    Movimento,
    NotaMotor,
    OperacaoMotor,
    PosicaoAtivo,
    TipoAjuste,
    TipoAtivo,
    TipoEvento,
    Visao,
)
from coir.domain.nota import NegocioBMF, NotaExtraida
from coir.domain.parametros import ReferenciasFiscais
from coir.domain.referencias import RE_SERIE_OPCAO, Referencias
from coir.engine.apuracao import apurar
from coir.engine.calendario import liquidacao_termo, vencimento_opcao
from coir.engine.futuros import movimentos_futuros
from coir.engine.movimentos import carteira, movimentar
from coir.erros import ReferenciaError

RE_PRAZO = re.compile(r"^(?P<mes>\d{2})/(?P<ano>\d{2})$")
RE_SEMANAL = re.compile(r"\dW(?P<semana>[1-5])$")  # strike + Wn; "ABEVW5" é put de novembro
# mercado do contrato (data/contratos_futuros.csv) -> linha da apuração; mercado sem linha aqui
# (ex.: juros) falha alto em vez de cair em "outros" calado
GRUPOS = {"INDICE": Grupo.INDICE, "DOLAR": Grupo.DOLAR, "BOI": Grupo.COMMODITY,
          "CAFE": Grupo.COMMODITY, "MILHO": Grupo.COMMODITY, "SOJA": Grupo.COMMODITY}  # fmt: skip
TIPOS_SEM_NEGOCIO = ("AJUPOS", "PERMAN")


class _Tickers:
    """(ticker, data) -> ativo, só leitura (a opção exercida já foi negociada antes)."""

    def __init__(self, sessao: Session) -> None:
        self._linhas = list(sessao.scalars(select(TickerHistorico)))
        self._por_ticker: dict[str, list[TickerHistorico]] = {}
        self._por_ativo: dict[int, list[TickerHistorico]] = {}
        self._por_papel = {
            p: a for p, a in sessao.execute(select(Operacao.papel, Operacao.ativo_id))
        }
        for t in self._linhas:
            self._por_ticker.setdefault(t.ticker, []).append(t)
            self._por_ativo.setdefault(t.ativo_id, []).append(t)

    @staticmethod
    def _cobre(t: TickerHistorico, dia: date) -> bool:
        return (t.inicio is None or t.inicio <= dia) and (t.fim is None or dia <= t.fim)

    def ativo(self, ticker: str, dia: date) -> int | None:
        return next((t.ativo_id for t in self._por_ticker.get(ticker, []) if self._cobre(t, dia)),
                    None)  # fmt: skip

    def papel(self, papel: str) -> int | None:
        """Série negociada com o código antigo (troca de ticker do ativo-objeto)."""
        return self._por_papel.get(papel)

    def ticker(self, ativo: int, dia: date) -> str:
        linhas = self._por_ativo.get(ativo, [])
        achado = next((t.ticker for t in linhas if self._cobre(t, dia)), None)
        if achado is None:
            raise ReferenciaError(f"ativo {ativo} sem ticker vigente em {dia}")
        return achado


def _vencimento(prazo: str, serie: str, refs: ReferenciasFiscais) -> date:
    m = RE_PRAZO.match(prazo)
    if m is None:
        raise ReferenciaError(f"opção {serie} sem prazo de vencimento impresso (lido {prazo!r})")
    semanal = RE_SEMANAL.search(serie)
    return vencimento_opcao(2000 + int(m["ano"]), int(m["mes"]), refs.calendario,
                            int(semanal["semana"]) if semanal else None)  # fmt: skip


def _nota_bovespa(
    registro: NotaCorretagem, nota: NotaExtraida, refs: ReferenciasFiscais, tickers: _Tickers
) -> tuple[OperacaoMotor, ...]:
    ops: list[OperacaoMotor] = []
    for o, negocio in zip(registro.operacoes, nota.negocios_bovespa, strict=True):
        opcao = o.mercado in ("CALL", "PUT")
        tipo = TipoAtivo.OPCAO if opcao else refs.tipos.tipo(o.papel)
        exercida = None
        if o.mercado.startswith("EXERC") and o.ativo_objeto:
            exercida = tickers.ativo(o.ativo_objeto, registro.data_pregao) or tickers.papel(
                o.ativo_objeto
            )
        isencao = refs.pme.get(o.papel)
        pme_ate = None
        if not opcao and o.cv == "C" and isencao and isencao.compra_elegivel(registro.data_pregao):
            pme_ate = isencao.vendas_ate
        liquidacao = None
        if o.mercado == "TERMO" and o.prazo_termo_dias is not None:
            liquidacao = liquidacao_termo(registro.data_pregao, o.prazo_termo_dias, refs.calendario)
        ops.append(
            OperacaoMotor(
                id=o.id, ordem=o.ordem, ativo=o.ativo_id, ticker=o.papel, tipo=tipo,
                mercado=o.mercado, cv=o.cv,  # type: ignore[arg-type]
                quantidade=o.quantidade, preco=o.preco, total=o.total,
                vencimento=_vencimento(negocio.prazo, o.papel, refs) if opcao else None,
                opcao=exercida, liquidacao_termo=liquidacao, pme_ate=pme_ate,
            )
        )  # fmt: skip
    return tuple(ops)


def _ajustes(nota: NotaExtraida, refs: Referencias) -> tuple[AjusteFuturo, ...]:
    saida: list[AjusteFuturo] = []
    for i, n in enumerate(nota.negocios_bmf):
        contrato = next(
            (
                c
                for c in refs.contratos
                if c.raiz == n.mercadoria[:3] and c.vigente(nota.data_pregao)
            ),
            None,
        )
        sem_negocio = any(t in n.tipo.upper() for t in TIPOS_SEM_NEGOCIO)
        if n.day_trade:
            tipo = TipoAjuste.DAY_TRADE
        elif "AJUPOS" in n.tipo.upper():
            tipo = TipoAjuste.AJUPOS
        elif sem_negocio:
            tipo = TipoAjuste.OUTRO
        else:
            tipo = TipoAjuste.NORMAL
        if contrato is None:
            if sem_negocio and n.valor == 0:
                continue  # taxa de permanência zerada de mercadoria sem cadastro
            raise ReferenciaError(f"contrato futuro {n.mercadoria} sem cadastro")
        if contrato.mercado not in GRUPOS:
            raise ReferenciaError(f"contrato futuro {n.mercadoria}: mercado {contrato.mercado} "
                                  "sem linha na apuração (ex.: juros)")  # fmt: skip
        saida.append(
            AjusteFuturo(
                ordem=i, mercadoria=n.mercadoria,
                grupo=GRUPOS[contrato.mercado], tipo=tipo,
                valor=_assinado(n), cv=n.cv, quantidade=abs(n.quantidade),
                vencimento=n.vencimento,
                volume=Decimal(0) if sem_negocio
                else abs(n.preco * n.quantidade * contrato.multiplicador),
            )
        )  # fmt: skip
    return tuple(saida)


def _assinado(n: NegocioBMF) -> Decimal:
    return n.valor if n.dc == "C" else -n.valor  # como na conciliação (sem indicador: débito)


def entradas(
    sessao: Session, refs: Referencias, fiscais: ReferenciasFiscais
) -> tuple[list[NotaMotor], list[Evento]]:
    tickers = _Tickers(sessao)
    notas: list[NotaMotor] = []
    registros = sessao.scalars(
        select(NotaCorretagem)
        .options(
            selectinload(NotaCorretagem.operacoes),
            selectinload(NotaCorretagem.conta).selectinload(Conta.corretora),
        )
        .order_by(NotaCorretagem.data_pregao, NotaCorretagem.id)
    )
    for r in registros:
        nota = NotaExtraida.model_validate_json(r.json_nota)
        bovespa = r.segmento == "bovespa"
        notas.append(
            NotaMotor(
                id=r.id, titular=r.conta.titular_id, conta=r.conta_id,
                corretora=r.conta.corretora.nome, numero=r.numero, data=r.data_pregao,
                custos=r.custos_total, irrf=r.irrf, irrf_day_trade=r.irrf_day_trade,
                operacoes=_nota_bovespa(r, nota, fiscais, tickers) if bovespa else (),
                ajustes=() if bovespa else _ajustes(nota, refs),
            )
        )  # fmt: skip
    eventos: list[Evento] = []
    for e in sessao.scalars(
        select(EventoCorporativo)
        .options(
            selectinload(EventoCorporativo.conta).selectinload(Conta.corretora),
            selectinload(EventoCorporativo.conta_destino).selectinload(Conta.corretora),
        )
        .order_by(EventoCorporativo.id)
    ):
        ticker = tickers.ticker(e.ativo_id, e.data)
        eventos.append(
            Evento(
                id=e.id, data=e.data, titular=e.conta.titular_id, conta=e.conta_id,
                corretora=e.conta.corretora.nome, ativo=e.ativo_id, ticker=ticker,
                tipo=TipoEvento(e.tipo),
                tipo_ativo=_tipo_evento(fiscais, ticker),
                cv=e.cv,  # type: ignore[arg-type]
                quantidade=e.quantidade, valor=e.valor, conta_destino=e.conta_destino_id,
                corretora_destino=e.conta_destino.corretora.nome if e.conta_destino else None,
                observacao=e.observacao,
            )
        )  # fmt: skip
    return notas, eventos


def _tipo_evento(fiscais: ReferenciasFiscais, ticker: str) -> TipoAtivo:
    if RE_SERIE_OPCAO.match(ticker):  # evento na série (ajuste da B3 após provento)
        return TipoAtivo.OPCAO
    return fiscais.tipos.tipo(ticker)


def carteira_banco(
    banco: Path, refs: Referencias, fiscais: ReferenciasFiscais, ate: date, visao: Visao
) -> tuple[list[PosicaoAtivo], dict[int, str]]:
    """Posições abertas no fim de `ate` na visão pedida, com o número de cada conta."""
    with conectar(banco) as fabrica, fabrica() as sessao:
        notas, eventos = entradas(sessao, refs, fiscais)
        contas = {c.id: c.numero for c in sessao.scalars(select(Conta))}
    return carteira(notas, eventos, fiscais.parametros, ate, visao), contas


def movimentos_banco(
    banco: Path, refs: Referencias, fiscais: ReferenciasFiscais, ate: date, visao: Visao
) -> list[Movimento]:
    """Movimentos (com futuros) na visão pedida; a apuração usa só a por titular."""
    with conectar(banco) as fabrica, fabrica() as sessao:
        notas, eventos = entradas(sessao, refs, fiscais)
    return movimentar(notas, eventos, fiscais.parametros, ate, visao) + movimentos_futuros(
        notas, ate
    )


def apurar_banco(
    banco: Path, refs: Referencias, fiscais: ReferenciasFiscais, ate: date
) -> Apuracao:
    with conectar(banco) as fabrica, fabrica() as sessao:
        notas, eventos = entradas(sessao, refs, fiscais)
    return apurar(notas, eventos, fiscais.parametros, fiscais.calendario, ate)
