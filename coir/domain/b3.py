"""Linhas das planilhas da Área do Investidor da B3 (fatos como a B3 os exporta).

As planilhas não trazem custos: servem para proventos, eventos de custódia e a conciliação com
as notas e com a posição. A instituição vem por extenso e com grafias variadas; o perfil da
corretora (XP, BTG...) sai de data/instituicoes_b3.csv.
"""

from datetime import date
from decimal import Decimal
from enum import StrEnum
from typing import Literal

from pydantic import BaseModel, ConfigDict


class TipoPlanilha(StrEnum):
    MOVIMENTACAO = "movimentacao"
    NEGOCIACAO = "negociacao"
    POSICAO = "posicao"  # relatório consolidado anual ou mensal (abas de posição)


class Modelo(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")


class LinhaMovimentacao(Modelo):
    data: date
    entrada: bool  # "Credito" (True) ou "Debito"
    movimentacao: str  # como impresso: "Dividendo", "Bonificação em Ativos", "Transferência"...
    produto: str  # "PETR4 - PETROBRAS" (o código vem antes do " - ")
    codigo: str  # código do produto (ticker, série, contrato); vazio se a B3 não traz
    instituicao: str
    quantidade: Decimal
    preco: Decimal | None
    valor: Decimal | None


class LinhaNegociacao(Modelo):
    data: date
    cv: Literal["C", "V"]
    mercado: str  # "Mercado à Vista", "Opção de Compra sobre Ações", "Futuro"...
    vencimento: date | None
    instituicao: str
    codigo: str
    quantidade: Decimal
    preco: Decimal
    valor: Decimal


class LinhaPosicao(Modelo):
    data: date  # data da posição (fim do mês do relatório consolidado)
    instituicao: str
    conta: str
    codigo: str
    isin: str
    quantidade: Decimal  # opção lançada, futuro vendido e ação tomada emprestada: negativa
    origem: Literal["acao", "emprestimo", "opcao", "futuro"] = "acao"
