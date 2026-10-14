"""Nota de corretagem extraída do PDF (valores como impressos, já em Decimal).

A nota é um fato: o que a corretora imprimiu. Normalização (ticker, mercado, categoria) e cálculo
(rateio, PM, apuração) vêm depois e nunca alteram estes objetos.
"""

from datetime import date
from decimal import Decimal
from enum import StrEnum
from typing import Literal, Self

from pydantic import BaseModel, ConfigDict, model_validator

DC = Literal["D", "C"]
CV = Literal["C", "V"]


class Segmento(StrEnum):
    BOVESPA = "bovespa"
    BMF = "bmf"


class CategoriaCusto(StrEnum):
    """Grupos de custo da nota (o rótulo impresso varia por corretora e ano)."""

    LIQUIDACAO = "liquidacao"
    REGISTRO = "registro"
    TERMO_OPCOES = "termo_opcoes"
    ANA = "ana"
    EMOLUMENTOS = "emolumentos"
    TRANSFERENCIA = "transferencia"
    CORRETAGEM = "corretagem"
    EXECUCAO = "execucao"
    CUSTODIA = "custodia"
    IMPOSTOS = "impostos"
    OUTROS = "outros"
    TAXAS_BMF = "taxas_bmf"


class Modelo(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")


class Montante(Modelo):
    """Valor impresso com o indicador D (débito) / C (crédito); dc None se a nota não imprime."""

    valor: Decimal
    dc: DC | None = None

    @property
    def assinado(self) -> Decimal:
        """Crédito positivo, débito negativo (sem indicador: como impresso)."""
        return -self.valor if self.dc == "D" else self.valor


class NegocioBovespa(Modelo):
    cv: CV
    mercado: str  # cru: VISTA, FRACIONARIO, TERMO, OPCAO DE COMPRA, EXERC OPC VENDA...
    prazo: str  # cru: "02/23" (opção), "90" (termo, dias) ou ""
    especificacao: str  # nome de pregão ("COPASA ON NM") ou código ("BBASN380 ON", "KLBNM199E")
    observacao: str  # marcas: D (day trade), #, @, H, 2, 8...
    quantidade: Decimal
    preco: Decimal
    valor: Decimal
    dc: DC
    lote_mil: bool = False  # cotação por lote de mil ("*"): valor = quantidade x preço / 1000

    @property
    def day_trade(self) -> bool:
        return "D" in self.observacao


class NegocioBMF(Modelo):
    cv: CV
    mercadoria: str  # sem espaços: WING19, WINZ18, CCMH24, BGIG22
    vencimento: date | None
    quantidade: Decimal  # AJUPOS de posição vendida vem negativa
    preco: Decimal
    tipo: str  # DAY TRADE, NORMAL, AJUPOS, TX. PERMANÊNCIA...
    valor: Decimal
    dc: DC | None
    taxa_operacional: Decimal

    @property
    def day_trade(self) -> bool:
        return self.tipo.replace(" ", "").upper() == "DAYTRADE"


GrupoCusto = Literal["cblc", "bolsa", "despesas", "bmf"]


class Custo(Modelo):
    rotulo: str
    categoria: CategoriaCusto
    grupo: GrupoCusto  # subtotal da nota em que o custo entra (Total CBLC, Bovespa/Soma, Despesas)
    valor: Decimal  # positivo = débito ao cliente


class ResumoBovespa(Modelo):
    valor_operacoes: Decimal
    compras: Decimal  # à vista + opções, como no "Resumo dos Negócios"
    vendas: Decimal
    termo: Decimal = Decimal(
        0
    )  # "Operações à termo": fora do líquido da nota (liquida no vencimento)
    valor_liquido_operacoes: Montante
    custos: tuple[Custo, ...]
    total_cblc: Montante
    total_bovespa: Montante
    total_despesas: Montante
    irrf: Decimal  # I.R.R.F. s/ operações (0,005%) impresso
    irrf_base: Decimal | None
    irrf_no_liquido: bool  # se o IRRF foi descontado do líquido desta nota
    irrf_day_trade: Decimal  # projeção do IRRF de day trade (1%)
    liquido: Montante
    data_liquidacao: date | None


class ResumoBMF(Modelo):
    valor_negocios: Montante
    ajuste_posicao: Montante
    ajuste_day_trade: Montante
    custos: tuple[Custo, ...]
    total_custos: Decimal
    impostos: Decimal
    irrf: Decimal
    irrf_day_trade: Decimal
    liquido: Montante


class NotaExtraida(Modelo):
    """Uma nota de corretagem (pode ocupar várias folhas). Chave: corretora, conta, número, data,
    segmento.
    """

    corretora: str
    conta: str
    numero: str
    data_pregao: date
    segmento: Segmento
    arquivo: str
    paginas: tuple[int, ...]  # 1..n no PDF
    negocios_bovespa: tuple[NegocioBovespa, ...] = ()
    negocios_bmf: tuple[NegocioBMF, ...] = ()
    resumo_bovespa: ResumoBovespa | None = None
    resumo_bmf: ResumoBMF | None = None
    avisos: tuple[str, ...] = ()

    @model_validator(mode="after")
    def _segmento_coerente(self) -> Self:
        if self.segmento is Segmento.BOVESPA:
            ok = self.resumo_bovespa is not None and not self.negocios_bmf and not self.resumo_bmf
        else:
            ok = (
                self.resumo_bmf is not None
                and not self.negocios_bovespa
                and not self.resumo_bovespa
            )
        if not ok:
            raise ValueError(f"nota {self.numero}: negócios/resumo incoerentes com o segmento")
        return self

    @property
    def chave(self) -> tuple[str, str, str, date, Segmento]:
        return (self.corretora, self.conta, self.numero, self.data_pregao, self.segmento)

    @property
    def custos_total(self) -> Decimal:
        """Custos da nota sem IRRF (o que o gabarito rateia nas operações)."""
        resumo = self.resumo_bovespa or self.resumo_bmf
        assert resumo is not None
        total = sum((c.valor for c in resumo.custos), Decimal(0))
        return total + (self.resumo_bmf.impostos if self.resumo_bmf else Decimal(0))

    @property
    def irrf_total(self) -> Decimal:
        resumo = self.resumo_bovespa or self.resumo_bmf
        assert resumo is not None
        return resumo.irrf + resumo.irrf_day_trade


class OperacaoExtraida(Modelo):
    """Negócio da nota já normalizado (ticker, mercado, categoria); valores ainda sem rateio de
    custos.
    """

    corretora: str
    conta: str
    nota: str
    data: date
    categoria: Literal["COMUM", "DAY_TRADE"]
    cv: CV
    papel: str  # ticker; no exercício, o ativo-objeto; em opções, o código da opção
    mercado: (
        str  # VISTA, FRACIONARIO, TERMO, CALL, PUT, EXERC CALL, EXERC PUT, INDICE, DOLAR, MILHO...
    )
    preco: Decimal
    quantidade: Decimal
    total: Decimal  # BM&F: preço x quantidade x multiplicador do contrato
    ativo_objeto: str | None  # opção: ativo-objeto; exercício: código da opção (sem o "E")
    prazo_termo_dias: int | None = None  # termo: prazo do contrato (dias), para a liquidação (F4)
