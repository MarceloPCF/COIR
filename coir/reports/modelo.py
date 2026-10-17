"""Modelo neutro de relatório: o mesmo objeto vira xlsx, csv e pdf e é exibido pela interface.

Um `Relatorio` tem cabeçalho (titular, período, data de geração) e seções; cada seção tem colunas
tipadas, linhas e uma linha de total opcional. Os valores ficam como Decimal, date, int ou texto;
a formatação (pt-BR) depende só do tipo da coluna.
"""

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import ROUND_HALF_EVEN, Decimal
from enum import StrEnum

from coir.erros import ExportacaoError


class TipoColuna(StrEnum):
    TEXTO = "texto"
    DATA = "data"
    INTEIRO = "inteiro"
    QUANTIDADE = "quantidade"  # até 8 casas, sem zeros à direita
    DINHEIRO = "dinheiro"  # 2 casas
    PM = "pm"  # 8 casas (preço médio, custo unitário)
    PRECO = "preco"  # preço negociado: de 2 a 8 casas, sem zeros à direita além da 2ª
    PERCENTUAL = "percentual"  # fração: 0,1234 = 12,34%


CASAS = {TipoColuna.DINHEIRO: 2, TipoColuna.PM: 8, TipoColuna.PRECO: 8, TipoColuna.QUANTIDADE: 8,
         TipoColuna.PERCENTUAL: 4, TipoColuna.INTEIRO: 0}  # fmt: skip
NUMERICAS = frozenset(CASAS)

Valor = str | Decimal | date | int | None


@dataclass(frozen=True)
class Coluna:
    nome: str
    tipo: TipoColuna = TipoColuna.TEXTO


@dataclass(frozen=True)
class Secao:
    titulo: str
    colunas: tuple[Coluna, ...]
    linhas: tuple[tuple[Valor, ...], ...]
    total: tuple[Valor, ...] | None = None  # mesma largura das colunas; rótulo na 1ª
    nota: str = ""  # texto curto abaixo da tabela (fonte legal, legenda)

    def __post_init__(self) -> None:
        largura = len(self.colunas)
        for i, linha in enumerate((*self.linhas, *((self.total,) if self.total else ()))):
            if len(linha) != largura:
                raise ExportacaoError(
                    f"seção {self.titulo!r}, linha {i + 1}: {len(linha)} valores para "
                    f"{largura} colunas"
                )


@dataclass(frozen=True)
class Cabecalho:
    titular: str  # nome
    cpf: str  # 11 dígitos
    periodo: str  # ex.: "Ano-calendário 2025" ou "Posição em 31/12/2025"
    gerado_em: datetime


@dataclass(frozen=True)
class Relatorio:
    chave: str  # nome curto para arquivo e aba (ex.: "resumo-fiscal")
    titulo: str
    cabecalho: Cabecalho
    secoes: tuple[Secao, ...]
    pdf: bool = False  # também sai em PDF
    notas: tuple[str, ...] = field(default=())  # observações gerais (fim do relatório)
    aba: str = ""  # nome curto da aba no xlsx (até 31 caracteres); vazio = título


# ---------------------------------------------------------------------------------- formatação
def arredondar(valor: Decimal, tipo: TipoColuna) -> Decimal:
    casas = CASAS.get(tipo, 8)
    return valor.quantize(Decimal(1).scaleb(-casas), rounding=ROUND_HALF_EVEN)


def numero_br(valor: Decimal, casas: int, milhar: bool = True) -> str:
    texto = f"{valor:,.{casas}f}" if milhar else f"{valor:.{casas}f}"
    return texto.replace(",", "_").replace(".", ",").replace("_", ".")


def formatar(valor: Valor, tipo: TipoColuna, milhar: bool = True) -> str:
    """Texto pt-BR do valor conforme o tipo da coluna (`milhar=False` no csv)."""
    if valor is None:
        return ""
    if isinstance(valor, str):
        return valor
    if isinstance(valor, datetime):
        return f"{valor:%d/%m/%Y %H:%M}"
    if isinstance(valor, date):
        return f"{valor:%d/%m/%Y}"
    numero = Decimal(valor)
    match tipo:
        case TipoColuna.PERCENTUAL:
            return numero_br(numero * 100, 2, milhar) + "%"
        case TipoColuna.QUANTIDADE | TipoColuna.PRECO:
            q = arredondar(numero, tipo).normalize()
            casas = max(0, -q.as_tuple().exponent)  # type: ignore[operator]
            if tipo is TipoColuna.PRECO:
                casas = max(casas, 2)
            return numero_br(q, casas, milhar)
        case _:
            return numero_br(arredondar(numero, tipo), CASAS.get(tipo, 2), milhar)


def cpf_formatado(cpf: str) -> str:
    d = re.sub(r"\D", "", cpf)
    return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}" if len(d) == 11 else cpf


def slug(texto: str) -> str:
    """Nome de arquivo/aba sem acento nem espaço."""
    tabela = str.maketrans("áàâãäéêèíìóôõòúùüçÁÀÂÃÉÊÍÓÔÕÚÇ", "aaaaaeeeiioooouuucAAAAEEIOOOUC")
    return re.sub(r"[^a-z0-9]+", "-", texto.translate(tabela).lower()).strip("-")
