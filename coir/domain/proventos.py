"""Proventos recebidos (dividendos, JCP, rendimentos de FII, reembolsos de BTC) classificados para
o IRPF a partir das linhas da movimentação da B3."""

from datetime import date
from decimal import Decimal
from enum import StrEnum

from pydantic import BaseModel, ConfigDict

from coir.domain.b3 import LinhaMovimentacao


class Modelo(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")


class TipoProvento(StrEnum):
    DIVIDENDO = "DIVIDENDO"  # isento (Lei 9.249/1995 art. 10)
    JCP = "JCP"  # tributação exclusiva na fonte (Lei 9.249/1995 art. 9º; LC 224/2025)
    RENDIMENTO_FII = "RENDIMENTO_FII"  # isento (Lei 11.033/2004 art. 3º III)
    REEMBOLSO_BTC = "REEMBOLSO_BTC"  # reembolso de proventos ao doador de ações emprestadas
    ATUALIZACAO = "ATUALIZACAO"  # atualização do provento pela Selic: tributação exclusiva
    RESTITUICAO_CAPITAL = "RESTITUICAO_CAPITAL"  # reduz o custo (IN RFB 1.585/2015 art. 58 §8º)
    FORA_DO_ESCOPO = "FORA_DO_ESCOPO"  # BDR, ETF, rendimento em ação: listado com aviso


class LinhaB3(Modelo):
    """Linha da movimentação já gravada (id para a origem, regra 12)."""

    id: int
    titular: int
    linha: LinhaMovimentacao
    chave: str = ""  # da linha em movimentacao_b3 (correções do usuário apontam para ela)
    # JCP: data do crédito (informada pelo usuário ou a data de aprovação do pacote da B3)
    credito: date | None = None
    credito_origem: str = ""  # "informada" ou "B3"


class Provento(Modelo):
    data: date  # pagamento (data da movimentação da B3)
    titular: int
    codigo: str
    empresa: str  # nome do produto na B3
    instituicao: str
    tipo: TipoProvento
    quantidade: Decimal
    liquido: Decimal  # creditado em conta
    bruto: Decimal  # JCP: líquido / (1 - alíquota); demais: igual ao líquido
    irrf: Decimal
    aliquota: Decimal
    ano_declaracao: int  # JCP creditado no ano anterior ao pagamento: o ano do crédito
    origem: str  # "mov:<id>"
    observacao: str = ""
    credito: date | None = None  # JCP: data do crédito, quando conhecida
    chave: str = ""  # da linha em movimentacao_b3 (tela de Proventos: corrigir o crédito)

    @property
    def a_receber_em_31_12(self) -> bool:
        """JCP declarado no ano do crédito e pago no ano seguinte (Bens e Direitos 99.07)."""
        return self.ano_declaracao < self.data.year
