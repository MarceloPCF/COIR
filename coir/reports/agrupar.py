"""Operações finalizadas por dia e ativo, em blocos por mês e mercado, para conferir com outro
controle sem somar à mão.

Os negócios do mesmo ativo no mesmo dia viram uma linha, por corretora, categoria e sentido; o mesmo
ativo em outro dia é outra linha. Só soma o que o motor já apurou: day trade e operação comum ficam
separados, porque são apurados à parte.
"""

from collections.abc import Iterable
from dataclasses import dataclass
from datetime import date
from decimal import Decimal

from coir.reports.modelo import Coluna, TipoColuna, Valor

T = TipoColuna
COLUNAS_POR_ATIVO = (
    Coluna("Ativo"),
    Coluna("Qtd.", T.QUANTIDADE),
    Coluna("Corretora"),
    Coluna("Categoria"),
    Coluna("C/V"),  # o que abriu a posição
    Coluna("Custo médio", T.PRECO),
    Coluna("C/V"),  # o que encerrou
    Coluna("Preço final", T.PRECO),
    Coluna("Data", T.DATA),
    Coluna("Lucratividade", T.PERCENTUAL),
    Coluna("Obs."),
    Coluna("Resultado", T.DINHEIRO),
)
COLUNA_ATIVO, COLUNA_DATA = 0, 8
ZERO = Decimal(0)
BOVESPA, BMF = "Bovespa", "BM&F"


@dataclass(frozen=True)
class ItemOperacao:
    """Uma operação; abertura e encerramento None: futuro, que só tem o resultado dos ajustes."""

    data: date
    ativo: str
    corretora: str
    categoria: str
    sentido: str  # "C/V": comprou e vendeu; "V/C": vendeu a descoberto e recomprou; "": pó
    quantidade: Decimal
    abertura: Decimal | None
    encerramento: Decimal | None
    resultado: Decimal
    obs: str = ""

    @property
    def mercado(self) -> str:
        return BMF if self.abertura is None else BOVESPA


def _unitario(valores: list[Decimal | None], quantidade: Decimal) -> Decimal | None:
    presentes = [v for v in valores if v is not None]
    return sum(presentes, ZERO) / quantidade if presentes and quantidade else None


def _linha(grupo: list[ItemOperacao]) -> tuple[Valor, ...]:
    primeiro = grupo[0]
    quantidade = sum((i.quantidade for i in grupo), ZERO)
    resultado = sum((i.resultado for i in grupo), ZERO)
    aberto = sum((i.abertura for i in grupo if i.abertura is not None), ZERO)
    abre, _, encerra = primeiro.sentido.partition("/")
    return (
        primeiro.ativo,
        quantidade,
        primeiro.corretora,
        primeiro.categoria,
        abre,
        _unitario([i.abertura for i in grupo], quantidade),
        encerra,
        _unitario([i.encerramento for i in grupo], quantidade),
        primeiro.data,
        resultado / aberto if aberto else None,
        "; ".join(dict.fromkeys(i.obs for i in grupo if i.obs)),
        resultado,
    )


def por_ativo(itens: Iterable[ItemOperacao]) -> list[tuple[Valor, ...]]:
    """Uma linha por dia, ativo, corretora, categoria e sentido, em ordem de data e ativo."""
    grupos: dict[tuple[date, str, str, str, str], list[ItemOperacao]] = {}
    for item in itens:
        chave = (item.data, item.ativo, item.corretora, item.categoria, item.sentido)
        grupos.setdefault(chave, []).append(item)
    return [_linha(grupos[chave]) for chave in sorted(grupos)]


def por_mes(itens: Iterable[ItemOperacao]) -> dict[tuple[int, int, str], list[tuple[Valor, ...]]]:
    """(ano, mês, mercado) -> as linhas de ``por_ativo``, em ordem de mês, com a Bovespa antes."""
    blocos: dict[tuple[int, int, str], list[ItemOperacao]] = {}
    for item in itens:
        blocos.setdefault((item.data.year, item.data.month, item.mercado), []).append(item)
    ordem = sorted(blocos, key=lambda b: (b[0], b[1], b[2] != BOVESPA))
    return {bloco: por_ativo(blocos[bloco]) for bloco in ordem}


def total_por_ativo(linhas: list[tuple[Valor, ...]]) -> tuple[Valor, ...]:
    """Rodapé: só o resultado (quantidade e preço de ativos diferentes não se somam)."""
    resultado = sum((r for x in linhas if isinstance(r := x[-1], Decimal)), ZERO)
    return ("Total", *[None] * (len(COLUNAS_POR_ATIVO) - 2), resultado)


def chave_da_linha(linha: tuple[Valor, ...]) -> str:
    """O que volta ao clicar numa linha: o dia e o ativo, para abrir as operações dele."""
    return f"{linha[COLUNA_DATA]}|{linha[COLUNA_ATIVO]}"
