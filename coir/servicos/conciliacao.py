"""Conciliação da posição calculada pelo COIR com a posição da B3 (relatório consolidado) numa
data — critério de saída da F7 (posição calculada = posição B3 no fim de cada ano).

Comparação por corretora e código de negociação, na visão por corretora (a B3 mostra a custódia de
cada instituição). A custódia da B3 muda na liquidação, o COIR conta no pregão: os negócios dos
últimos pregões que liquidam depois da data saem da comparação (ações, units, BDR e ETF em D+2
desde 27/05/2019, antes D+3; as opções a B3 já mostra no dia do pregão). Na B3, as ações
emprestadas (BTC, doador) continuam do titular e somam à posição; opção lançada entra negativa.
Diferença ganha uma causa provável: os eventos que a própria
movimentação da B3 registrou para aquele código desde a posição anterior (evento corporativo ainda
não lançado no COIR é a causa mais comum).
"""

import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path

from sqlalchemy import select
from sqlalchemy.orm import Session

from coir.db.modelos import (
    Conta,
    Corretora,
    MovimentacaoB3,
    NotaCorretagem,
    Operacao,
    PosicaoB3,
)
from coir.db.sessao import conectar
from coir.domain.fiscal import PosicaoAtivo, TipoAtivo, Visao
from coir.domain.parametros import ReferenciasFiscais
from coir.domain.referencias import Referencias
from coir.engine.calendario import UM_DIA, dia_de_pregao
from coir.servicos.apuracao import carteira_banco

ZERO = Decimal(0)
D2_DESDE = date(2019, 5, 27)  # B3: ciclo de liquidação de ações passou de D+3 para D+2
JANELA = timedelta(days=10)
RE_SERIE = re.compile(r"^[A-Z0-9]{4}[A-X]\d")
RE_CODIGO = re.compile(r"^[A-Z0-9]{4}(?:\d{1,2}|[A-X]\d{1,4}(?:W\d)?)$")  # ação/unit/BDR ou série
EVENTOS_B3 = (
    "Desdobro",
    "Bonificação em Ativos",
    "Incorporação",
    "Grupamento",
    "Restituição de Capital",
    "Direitos de Subscrição - Exercido",
    "Recibo de Subscrição",
    "Leilão de Fração",
    "Fração em Ativos",
)


@dataclass(frozen=True)
class LinhaConciliacao:
    corretora: str
    codigo: str
    b3: Decimal
    coir: Decimal
    causa: str = ""

    @property
    def diferenca(self) -> Decimal:
        return self.coir - self.b3

    @property
    def confere(self) -> bool:
        return self.diferenca == 0


@dataclass(frozen=True)
class Conciliacao:
    data: date
    linhas: tuple[LinhaConciliacao, ...]
    avisos: tuple[str, ...] = field(default=())

    @property
    def divergentes(self) -> tuple[LinhaConciliacao, ...]:
        return tuple(x for x in self.linhas if not x.confere)


def datas_b3(banco: Path, titular: int) -> list[date]:
    """Datas com posição da B3 importada, da mais antiga à mais recente."""
    with conectar(banco) as fabrica, fabrica() as s:
        consulta = select(PosicaoB3.data).where(PosicaoB3.titular_id == titular).distinct()
        return sorted(s.scalars(consulta))


def conciliar(
    banco: Path,
    refs: Referencias,
    fiscais: ReferenciasFiscais,
    titular: int,
    dia: date,
    anterior: date | None = None,
    posicoes: list[PosicaoAtivo] | None = None,
) -> Conciliacao:
    """B3 x COIR na data `dia`; `anterior` (a posição B3 anterior) limita a busca de causas.
    `posicoes`: a carteira do COIR na data por corretora, já calculada (várias datas de uma vez:
    ``carteiras_banco``); sem ela, calcula aqui."""
    avisos: list[str] = []
    b3: dict[tuple[str, str], Decimal] = defaultdict(Decimal)
    with conectar(banco) as fabrica, fabrica() as s:
        for p in s.scalars(
            select(PosicaoB3).where(PosicaoB3.titular_id == titular, PosicaoB3.data == dia)
        ):
            if p.origem == "futuro":
                continue  # futuros: resultado por ajustes, sem posição de custódia (tela 9)
            if not RE_CODIGO.match(p.codigo):
                avisos.append(
                    f"{p.codigo}: código fora de ação/opção (renda fixa, recibo?), não conciliado"
                )
                continue
            corretora = refs.corretora_b3(p.instituicao)
            if corretora is None:
                avisos.append(
                    f"instituição sem cadastro em data/instituicoes_b3.csv: {p.instituicao}"
                )
                corretora = p.instituicao
            # o COIR guarda a posição com o código atual (troca de ticker): o da B3 também
            b3[(corretora, refs.ticker_na_data(p.codigo, dia))] += p.quantidade
        eventos = list(
            s.scalars(
                select(MovimentacaoB3)
                .where(
                    MovimentacaoB3.titular_id == titular,
                    MovimentacaoB3.data <= dia,
                    MovimentacaoB3.movimentacao.in_(EVENTOS_B3),
                )
                .order_by(MovimentacaoB3.data)
            )
        )
        a_liquidar = _a_liquidar(s, titular, dia, fiscais)
    if posicoes is None:
        posicoes, _ = carteira_banco(banco, refs, fiscais, dia, Visao.CORRETORA)
    coir: dict[tuple[str, str], Decimal] = defaultdict(Decimal)
    for p in posicoes:
        if p.titular == titular and p.quantidade and p.tipo is not TipoAtivo.FUTURO:
            coir[(p.corretora or "", p.ticker)] += p.quantidade
    for chave, quantidade in a_liquidar.items():
        coir[chave] -= quantidade  # ainda não liquidado: fora da custódia da B3 na data
    linhas: list[LinhaConciliacao] = []
    for chave in sorted(set(b3) | set(coir)):
        q_b3, q_coir = b3.get(chave, ZERO), coir.get(chave, ZERO)
        causa = "" if q_b3 == q_coir else _causa(chave[1], eventos, anterior, q_b3, q_coir)
        linhas.append(LinhaConciliacao(chave[0], chave[1], q_b3, q_coir, causa))
    return Conciliacao(dia, tuple(linhas), tuple(dict.fromkeys(avisos)))


def _causa(
    codigo: str, eventos: list[MovimentacaoB3], anterior: date | None, b3: Decimal, coir: Decimal
) -> str:
    raiz = codigo[:4] if RE_SERIE.match(codigo) else codigo  # série: eventos do ativo-objeto
    achados = [e for e in eventos if (e.codigo == codigo or e.codigo.startswith(raiz))
               and (anterior is None or e.data > anterior)]  # fmt: skip
    if achados:
        partes = [
            f"{e.movimentacao} em {e.data:%d/%m/%Y} ({'+' if e.entrada else '-'}"
            f"{e.quantidade.normalize()})"
            for e in achados[-3:]
        ]
        return "a B3 registrou: " + "; ".join(partes) + " — confira se o evento está lançado"
    if b3 == 0:
        return "só no COIR: venda, transferência ou evento sem nota/lançamento?"
    if coir == 0:
        return "só na B3: compra sem nota importada, subscrição ou evento não lançado?"
    return "quantidades diferentes: confira notas e eventos do período"


def _liquidacao(pregao: date, mercado: str, fiscais: ReferenciasFiscais) -> date | None:
    """Data em que a B3 move a custódia (None: mercado sem regra aqui)."""
    if mercado in ("VISTA", "FRACIONARIO"):
        prazo = 2 if pregao >= D2_DESDE else 3
    else:
        return None
    dia = pregao
    while prazo:
        dia += UM_DIA
        if dia_de_pregao(dia, fiscais.calendario):
            prazo -= 1
    return dia


def _a_liquidar(
    s: Session, titular: int, dia: date, fiscais: ReferenciasFiscais
) -> dict[tuple[str, str], Decimal]:
    """Quantidade (comprada +, vendida -) negociada até `dia` e liquidada depois dele."""
    saida: dict[tuple[str, str], Decimal] = defaultdict(Decimal)
    linhas = s.execute(
        select(
            NotaCorretagem.data_pregao,
            Corretora.nome,
            Operacao.papel,
            Operacao.mercado,
            Operacao.cv,
            Operacao.quantidade,
        )
        .join(Operacao, Operacao.nota_id == NotaCorretagem.id)
        .join(Conta, Conta.id == NotaCorretagem.conta_id)
        .join(Corretora, Corretora.id == Conta.corretora_id)
        .where(
            Conta.titular_id == titular,
            NotaCorretagem.data_pregao <= dia,
            NotaCorretagem.data_pregao > dia - JANELA,
        )
    ).all()
    for pregao, corretora, papel, mercado, cv, quantidade in linhas:
        liquida = _liquidacao(pregao, mercado, fiscais)
        if liquida is not None and liquida > dia:
            saida[(corretora, papel)] += quantidade if cv == "C" else -quantidade
    return saida
