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
    NaturezaEvento,
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
from coir.engine.apuracao import conta_na_isencao as conta_na_isencao  # para a tela
from coir.engine.calendario import liquidacao_termo, vencimento_opcao
from coir.engine.futuros import movimentos_futuros
from coir.engine.movimentos import carteira, movimentar
from coir.erros import ReferenciaError
from coir.parsing import futuros
from coir.servicos import outros_custos, saldos_iniciais

RE_PRAZO = re.compile(r"^(?P<mes>\d{2})/(?P<ano>\d{2})$")
RE_SEMANAL = re.compile(r"\dW(?P<semana>[1-5])$")  # strike + Wn; "ABEVW5" é put de novembro
# linha da apuração de cada contrato: coluna grupo de data/contratos_futuros.csv; grupo sem linha
# falha alto em vez de cair em "outros" calado
GRUPOS_FUTUROS = {Grupo.INDICE, Grupo.DOLAR, Grupo.JUROS, Grupo.COMMODITY}


class Tickers:
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


def _objeto_tipo(serie: str, publicas: Referencias, refs: ReferenciasFiscais) -> TipoAtivo | None:
    """Tipo do ativo-objeto de uma série, pela raiz de 4 letras (opcoes_raiz*.csv); None se a raiz
    não está nas tabelas (o motor a mantém na linha "opções - ações")."""
    objeto = publicas.opcoes_raiz.get(serie[:4])
    if objeto is None:
        return None
    try:
        return refs.tipos.tipo(objeto)
    except ReferenciaError:
        return None


def _nota_bovespa(
    registro: NotaCorretagem, nota: NotaExtraida, refs: ReferenciasFiscais, tickers: Tickers,
    publicas: Referencias,
) -> tuple[OperacaoMotor, ...]:  # fmt: skip
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
                marcada_dt=negocio.day_trade,
                objeto_tipo=_objeto_tipo(o.papel, publicas, refs) if opcao else None,
            )
        )  # fmt: skip
    return tuple(ops)


def _ajustes(nota: NotaExtraida, refs: Referencias) -> tuple[AjusteFuturo, ...]:
    conversao = futuros.conversao_do_volume(nota, refs)  # volume em R$ (câmbio implícito)
    saida: list[AjusteFuturo] = []
    for i, n in enumerate(nota.negocios_bmf):
        tipo = futuros.tipo(n)
        contrato = futuros.contrato(n, refs, nota.data_pregao)
        if contrato is None:
            if tipo is TipoAjuste.OUTRO and n.valor == 0:
                continue  # taxa de permanência zerada de mercadoria sem cadastro
            raise ReferenciaError(f"contrato futuro {n.mercadoria} sem cadastro")
        grupo = next((g for g in GRUPOS_FUTUROS if g.value == contrato.grupo), None)
        if grupo is None:
            raise ReferenciaError(f"contrato futuro {n.mercadoria}: grupo {contrato.grupo} "
                                  "sem linha na apuração")  # fmt: skip
        sem_negocio = tipo in (TipoAjuste.AJUPOS, TipoAjuste.OUTRO)
        volume = abs(n.preco * n.quantidade * contrato.multiplicador)
        saida.append(
            AjusteFuturo(
                ordem=i, mercadoria=n.mercadoria, grupo=grupo, tipo=tipo,
                valor=_assinado(n), cv=n.cv, quantidade=abs(n.quantidade),
                vencimento=n.vencimento,
                volume=Decimal(0) if sem_negocio else volume * conversao.get(n.mercadoria, 1),
                entrega_fisica=contrato.liquidacao == "FISICA",
            )
        )  # fmt: skip
    return tuple(saida)


def _assinado(n: NegocioBMF) -> Decimal:
    return n.valor if n.dc == "C" else -n.valor  # como na conciliação (sem indicador: débito)


def entradas(
    sessao: Session, refs: Referencias, fiscais: ReferenciasFiscais
) -> tuple[list[NotaMotor], list[Evento]]:
    tickers = Tickers(sessao)
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
                operacoes=_nota_bovespa(r, nota, fiscais, tickers, refs) if bovespa else (),
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
        destino = None if e.ativo_destino_id is None else tickers.ticker(e.ativo_destino_id,
                                                                         e.data)  # fmt: skip
        eventos.append(
            Evento(
                id=e.id, data=e.data, titular=e.conta.titular_id, conta=e.conta_id,
                corretora=e.conta.corretora.nome, ativo=e.ativo_id, ticker=ticker,
                tipo=TipoEvento(e.tipo),
                tipo_ativo=_tipo_evento(fiscais, ticker),
                cv=e.cv,  # type: ignore[arg-type]
                quantidade=e.quantidade, valor=e.valor, conta_destino=e.conta_destino_id,
                corretora_destino=e.conta_destino.corretora.nome if e.conta_destino else None,
                ativo_destino=e.ativo_destino_id, ticker_destino=destino,
                tipo_ativo_destino=None if destino is None else _tipo_evento(fiscais, destino),
                quantidade_destino=e.quantidade_destino,
                natureza=None if e.natureza is None else NaturezaEvento(e.natureza),
                percentual_custo=e.percentual_custo, observacao=e.observacao,
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


def carteiras_banco(
    banco: Path, refs: Referencias, fiscais: ReferenciasFiscais, datas: list[date], visao: Visao
) -> dict[date, list[PosicaoAtivo]]:
    """Posições em várias datas lendo o banco uma vez só (conciliação com a B3)."""
    with conectar(banco) as fabrica, fabrica() as sessao:
        notas, eventos = entradas(sessao, refs, fiscais)
    return {dia: carteira(notas, eventos, fiscais.parametros, dia, visao) for dia in datas}


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
        custos = outros_custos.do_banco(sessao)
        saldos = saldos_iniciais.do_banco(sessao)
    return apurar(notas, eventos, fiscais.parametros, fiscais.calendario, ate, custos, saldos)
