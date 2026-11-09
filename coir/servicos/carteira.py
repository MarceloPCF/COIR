"""Carteira para a tela: posições numa data, na visão pedida, e a trilha do preço médio de um
ativo (os movimentos do motor, com a regra de cada passo e a nota de origem)."""

import re
from collections import defaultdict
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path

from sqlalchemy import select

from coir.db.modelos import Operacao
from coir.db.sessao import conectar
from coir.domain.fiscal import Movimento, PosicaoAtivo, TermoAberto, TipoAtivo, Visao

TIPOS = {
    TipoAtivo.ACAO: "Ação", TipoAtivo.UNIT: "Unit", TipoAtivo.ETF: "ETF",
    TipoAtivo.ETF_RF: "ETF de renda fixa", TipoAtivo.BDR: "BDR", TipoAtivo.FII: "FII/Fiagro",
    TipoAtivo.OPCAO: "Opção", TipoAtivo.FUTURO: "Futuro",
    TipoAtivo.DIREITO: "Direito/recibo/bônus de subscrição",
    TipoAtivo.FUNDO_INFRA: "FI-Infra/FIP-IE",
    TipoAtivo.SEM_TIPO: "Sem tipo",
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
    liquidacao: date | None = None  # compra a termo ainda não liquidada (fora do PM)
    nota: int | None = None  # termo: a nota da compra, se for uma só


def posicoes(
    carteira: list[PosicaoAtivo],
    contas: dict[int, str],
    titular: int | None,
    termos: Iterable[TermoAberto] = (),
    visao: Visao = Visao.TITULAR,
) -> list[LinhaPosicao]:
    """As posições abertas do titular, por tipo e ticker. A compra a termo ainda não liquidada
    vem numa linha própria logo abaixo do ativo, como na B3: não está na posição nem no PM."""
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
    ordem = {id(x): x.tipo for x in linhas}
    grupos: dict[tuple[object, ...], list[TermoAberto]] = defaultdict(list)
    for t in termos:
        if titular is None or t.titular == titular:
            onde = {Visao.TITULAR: (t.titular,), Visao.CORRETORA: (t.titular, t.corretora),
                    Visao.CONTA: (t.titular, t.conta)}[visao]  # fmt: skip
            grupos[(*onde, t.ativo, t.liquidacao)].append(t)
    for grupo in grupos.values():
        t = grupo[0]
        quantidade = sum((x.quantidade for x in grupo), Decimal(0))
        custo = sum((x.custo for x in grupo), Decimal(0))
        conta = t.conta if visao is Visao.CONTA else None
        linha = LinhaPosicao(
            ativo=t.ativo,
            ticker=t.ticker,
            tipo=f"Termo · liquida em {t.liquidacao:%d/%m/%Y}",
            corretora=t.corretora if visao is not Visao.TITULAR else "",
            conta=conta,
            numero_conta=contas.get(conta, "") if conta is not None else "",
            vencimento=None,
            quantidade=quantidade,
            pm=custo / quantidade,
            custo=custo,
            liquidacao=t.liquidacao,
            nota=t.nota if len({x.nota for x in grupo}) == 1 else None,
        )
        linhas.append(linha)
        ordem[id(linha)] = TIPOS[t.tipo]
    return sorted(linhas, key=lambda x: (ordem[id(x)], x.ticker, x.corretora, x.numero_conta,
                                         x.liquidacao or date.min))  # fmt: skip


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
