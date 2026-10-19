"""Carteira para a tela: posições numa data, na visão pedida, e a trilha do preço médio de um
ativo (os movimentos do motor, com a regra de cada passo e a nota de origem)."""

import re
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path

from sqlalchemy import select

from coir.db.modelos import Operacao
from coir.db.sessao import conectar
from coir.domain.fiscal import Movimento, PosicaoAtivo, TipoAtivo

TIPOS = {
    TipoAtivo.ACAO: "Ação", TipoAtivo.UNIT: "Unit", TipoAtivo.ETF: "ETF",
    TipoAtivo.ETF_RF: "ETF de renda fixa", TipoAtivo.BDR: "BDR", TipoAtivo.FII: "FII/Fiagro",
    TipoAtivo.OPCAO: "Opção", TipoAtivo.FUTURO: "Futuro",
}  # fmt: skip
RE_ORIGEM = re.compile(r"^(op|termo|nota):(\d+)")


@dataclass(frozen=True)
class LinhaPosicao:
    ativo: int
    ticker: str
    tipo: str
    corretora: str  # vazio na visão por titular
    conta: int | None
    numero_conta: str
    vencimento: date | None
    quantidade: Decimal
    pm: Decimal
    custo: Decimal


def posicoes(
    carteira: list[PosicaoAtivo], contas: dict[int, str], titular: int | None
) -> list[LinhaPosicao]:
    """As posições abertas do titular, por tipo e ticker."""
    linhas = [
        LinhaPosicao(
            ativo=p.ativo,
            ticker=p.ticker,
            tipo=TIPOS[p.tipo],
            corretora=p.corretora or "",
            conta=p.conta,
            numero_conta=contas.get(p.conta, "") if p.conta is not None else "",
            vencimento=p.vencimento,
            quantidade=p.quantidade,
            pm=p.pm,
            custo=p.custo,
        )
        for p in carteira
        if (titular is None or p.titular == titular) and p.quantidade != 0
    ]
    return sorted(linhas, key=lambda x: (x.tipo, x.ticker, x.corretora, x.numero_conta))


def trilha(movimentos: list[Movimento], titular: int | None, posicao: LinhaPosicao,
           ate: date) -> list[Movimento]:  # fmt: skip
    """Os passos do ativo até a data, na mesma visão da posição (conta ou corretora, se houver)."""
    return [
        m
        for m in movimentos
        if (titular is None or m.titular == titular)
        and m.ativo == posicao.ativo
        and m.data <= ate
        and (posicao.conta is None or m.conta == posicao.conta)
        and (not posicao.corretora or m.corretora == posicao.corretora)
    ]


def notas_das_operacoes(banco: Path) -> dict[int, int]:
    """Operação -> nota, para abrir a nota de um movimento."""
    with conectar(banco) as fabrica, fabrica() as s:
        return {op: nota for op, nota in s.execute(select(Operacao.id, Operacao.nota_id))}


def nota_de(origem: str, mapa: dict[int, int]) -> int | None:
    """A nota de origem de um movimento ("op:12,13", "termo:7", "nota:4"); eventos não têm."""
    m = RE_ORIGEM.match(origem)
    if m is None:
        return None
    numero = int(m[2])
    return numero if m[1] == "nota" else mapa.get(numero)
