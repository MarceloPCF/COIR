"""Opções e futuros para as telas 8 e 9, a partir dos movimentos do motor (nada recalculado).

Opções: cada ciclo de uma série vai da abertura até a posição voltar a zero (o mesmo código pode
voltar em outro ano); a situação sai do último passo: aberta, vencida (pó), exercida ou
encerrada. Futuros: o resultado realizado por mês e contrato, os contratos em aberto (ajustes
acumulados) e as linhas de ajuste de cada contrato nas notas.
"""

from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path

from sqlalchemy import select

from coir.db.modelos import SerieOpcao
from coir.db.sessao import conectar
from coir.domain.fiscal import (
    ZERO,
    AjusteFuturo,
    Categoria,
    Movimento,
    PosicaoAtivo,
    PosicaoFuturo,
    TipoAtivo,
)
from coir.domain.parametros import ReferenciasFiscais
from coir.domain.referencias import Referencias
from coir.engine.futuros import movimentos_futuros
from coir.servicos.apuracao import entradas


@dataclass(frozen=True)
class CicloOpcao:
    ticker: str
    objeto: str  # ativo-objeto (4 letras + classe da abertura da série)
    call: bool
    abertura: date
    fim: date
    vencimento: date | None
    situacao: str  # aberta, vencida (pó), exercida, encerrada
    posicao: Decimal  # posição atual (aberta) ou a máxima do ciclo, com sinal
    premio: Decimal  # vendas - compras (prêmio líquido recebido > 0)
    resultado: Decimal
    passos: tuple[Movimento, ...]


def objetos(banco: Path, refs: Referencias) -> dict[str, str]:
    """Série -> ativo-objeto, pela classe vista na abertura (tabela serie_opcao)."""
    with conectar(banco) as fabrica, fabrica() as s:
        return {
            r.serie: r.serie[:4] + refs.classes[r.classe]
            for r in s.scalars(select(SerieOpcao))
            if r.classe in refs.classes
        }


def ciclos_opcoes(
    movimentos: list[Movimento], carteira: list[PosicaoAtivo], titular: int,
    objetos: dict[str, str], refs: Referencias,
) -> list[CicloOpcao]:  # fmt: skip
    vencimentos = {
        p.ticker: p.vencimento
        for p in carteira
        if p.titular == titular and p.tipo is TipoAtivo.OPCAO
    }
    por_serie: dict[str, list[Movimento]] = defaultdict(list)
    for m in movimentos:
        if m.titular == titular and m.tipo is TipoAtivo.OPCAO and m.ativo is not None:
            por_serie[m.ticker].append(m)
    ciclos: list[CicloOpcao] = []
    for ticker, movs in por_serie.items():
        objeto = objetos.get(ticker) or refs.opcoes_raiz.get(ticker[:4]) or ticker[:4]
        atual: list[Movimento] = []
        for m in sorted(movs, key=lambda m: m.data):
            atual.append(m)
            if m.quantidade_apos == 0:
                ciclos.append(_ciclo(ticker, objeto, atual, None))
                atual = []
        if atual:
            ciclos.append(_ciclo(ticker, objeto, atual, vencimentos.get(ticker)))
    return sorted(ciclos, key=lambda c: (c.fim, c.ticker), reverse=True)


def _ciclo(
    ticker: str, objeto: str, passos: list[Movimento], vencimento: date | None
) -> CicloOpcao:
    ultimo = passos[-1]
    premio = sum(((m.valor if m.cv == "V" else -m.valor) for m in passos if m.cv), ZERO)
    maior = max((m.quantidade_apos for m in passos), key=abs)
    if ultimo.quantidade_apos:
        situacao, posicao = "aberta", ultimo.quantidade_apos
    elif ultimo.origem == "vencimento":
        situacao, posicao, vencimento = "vencida (pó)", maior, ultimo.data
    elif "exercida" in ultimo.regra or "exercício" in ultimo.regra:
        situacao, posicao = "exercida", maior
    else:
        situacao, posicao = "encerrada", maior
    return CicloOpcao(
        ticker=ticker,
        objeto=objeto,
        call=ticker[4] <= "L",  # letra do mês: A-L call, M-X put
        abertura=passos[0].data,
        fim=ultimo.data,
        vencimento=vencimento,
        situacao=situacao,
        posicao=posicao,
        premio=premio,
        resultado=sum((m.resultado for m in passos), ZERO),
        passos=tuple(passos),
    )


@dataclass(frozen=True)
class LinhaAjuste:
    data: date
    nota: int
    numero: str
    mercadoria: str
    tipo: str
    cv: str
    quantidade: Decimal
    valor: Decimal


@dataclass(frozen=True)
class Futuros:
    abertas: tuple[PosicaoFuturo, ...]
    realizados: tuple[Movimento, ...]  # resultados de futuros do titular
    ajustes: tuple[LinhaAjuste, ...]
    avisos: tuple[str, ...] = ()


def futuros_banco(
    banco: Path, refs: Referencias, fiscais: ReferenciasFiscais, titular: int, ate: date
) -> Futuros:
    with conectar(banco) as fabrica, fabrica() as s:
        notas, _ = entradas(s, refs, fiscais)
    do_titular = [n for n in notas if n.titular == titular]
    abertas: list[PosicaoFuturo] = []
    avisos: list[str] = []
    realizados = movimentos_futuros(do_titular, ate, avisos, abertas)
    ajustes = [
        LinhaAjuste(
            n.data, n.id, n.numero, a.mercadoria, _tipo(a), a.cv or "", a.quantidade, a.valor
        )
        for n in do_titular
        if n.data <= ate
        for a in n.ajustes
    ]
    return Futuros(tuple(abertas), tuple(realizados), tuple(ajustes), tuple(dict.fromkeys(avisos)))


def _tipo(a: AjusteFuturo) -> str:
    return {
        "NORMAL": "negócio",
        "AJUPOS": "ajuste de posição",
        "DAY_TRADE": "day trade",
        "OUTRO": "taxa",
    }[a.tipo.value]


def resultado_mensal(
    realizados: tuple[Movimento, ...], ano: int
) -> list[tuple[int, str, Decimal, Decimal, Decimal]]:
    """(mês, contrato, day trade, comum, custos) do ano, sem as linhas zeradas."""
    soma: dict[tuple[int, str], list[Decimal]] = defaultdict(lambda: [ZERO, ZERO, ZERO])
    for m in realizados:
        if m.data.year != ano:
            continue
        linha = soma[(m.data.month, m.ticker)]
        linha[0 if m.categoria is Categoria.DAY_TRADE else 1] += m.resultado
        linha[2] += m.custos
    return [
        (mes, contrato, v[0], v[1], v[2])
        for (mes, contrato), v in sorted(soma.items())
        if any(v)  # taxa de permanência zerada (ex.: CCMPPPP) não vira linha
    ]
