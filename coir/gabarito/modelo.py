"""Registros do gabarito (pseudonimizados). Dinheiro e quantidades em Decimal."""

from datetime import date
from decimal import Decimal
from enum import StrEnum
from typing import Literal

from pydantic import BaseModel, ConfigDict


class Categoria(StrEnum):
    COMUM = "COMUM"
    DAY_TRADE = "DAY_TRADE"


class Registro(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")


class OperacaoGabarito(Registro):
    """Linha de Normais_Dados / DayTrade_Dados que veio de uma nota de corretagem."""

    categoria: Categoria
    linha: int
    titular: str
    corretora: str
    conta: str
    nota: str
    data: date
    cv: Literal["C", "V"]
    papel: str
    mercado: str
    preco: Decimal
    quantidade: Decimal
    total: Decimal
    custos: Decimal
    pm: Decimal | None
    irrf: Decimal
    prazo: date | None
    ativo_objeto: str | None


class EventoGabarito(Registro):
    """Lançamento manual sem nota (bonificação, desdobramento, grupamento, subscrição...)."""

    categoria: Categoria
    linha: int
    titular: str
    corretora: str
    conta: str
    data: date
    cv: Literal["C", "V"]
    papel: str
    tipo: str
    preco: Decimal | None
    quantidade: Decimal | None
    total: Decimal | None


class PosicaoGabarito(Registro):
    """Linha calculada de Normais_Op / DayTrade_Op: estado da posição após cada operação.

    ``ano_competencia``/``mes_competencia`` são o mês em que a planilha apura o resultado (difere
    da data do pregão em opções lançadas, apuradas no encerramento). ``isento`` é a marca da
    planilha para vendas de mês isento (excluídas da apuração comum).
    """

    categoria: Categoria
    ordem: int
    corretora: str
    conta: str
    papel: str
    mercado: str
    cv: Literal["C", "V"]
    data: date
    nota: str | None
    quantidade: Decimal
    pm_operacao: Decimal | None
    pm_apos: Decimal | None
    quantidade_apos: Decimal | None
    custos: Decimal | None
    irrf: Decimal | None
    resultado: Decimal | None
    ano_competencia: int
    mes_competencia: int
    isento: bool


class ApuracaoMensalGabarito(Registro):
    categoria: Categoria
    ano: int
    mes: int
    resultado: Decimal
    a_vista: Decimal
    opcoes: Decimal
    indice: Decimal
    dolar: Decimal
    commodity: Decimal
    irpf: Decimal
    irrf: Decimal
    a_pagar: Decimal
    acumulado: Decimal
    base_calculo: Decimal


class DarfGabarito(Registro):
    ano: int
    mes: int
    apuracao: date | None
    vencimento: date | None
    pagamento: date | None
    codigo: str | None
    valor: Decimal
    multa: Decimal | None
    juros: Decimal | None
    total: Decimal | None


class IsentoMensal(Registro):
    ano: int
    mes: int
    valor: Decimal


class CustodiaInicialGabarito(Registro):
    linha: int
    corretora: str
    data: date
    cv: str
    papel: str
    operacao: str
    preco_medio: Decimal
    quantidade: Decimal
    total: Decimal | None
    prejuizo_acumulado: Decimal | None


class Gabarito(BaseModel):
    """Conjunto completo extraído da planilha."""

    operacoes: list[OperacaoGabarito]
    eventos: list[EventoGabarito]
    eventos_cadastro: list[EventoGabarito]
    posicoes: list[PosicaoGabarito]
    apuracao: list[ApuracaoMensalGabarito]
    darfs: list[DarfGabarito]
    isentos: list[IsentoMensal]
    custodia_inicial: list[CustodiaInicialGabarito]
    avisos: list[str]
