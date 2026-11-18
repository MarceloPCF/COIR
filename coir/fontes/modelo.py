"""Registros das fontes de referência, por conceito e com a origem de cada um.

Os campos guardam o texto como está na tabela (datas ISO, números com ponto): quem usa converte,
como antes. Cada registro sabe de que fonte veio (``origem``), e o serviço não precisa saber de
qual arquivo ele saiu.
"""

from collections.abc import Mapping
from dataclasses import dataclass, fields
from datetime import date
from decimal import Decimal

# origens: curadoria do projeto, dados abertos da CVM (formulário de referência e avisos), do
# Banco Central e os arquivos públicos da bolsa
CURADORIA = "curadoria"
CVM_FRE = "cvm_fre"
CVM_IPE = "cvm_ipe"
CVM_FCA = "cvm_fca"  # formulário cadastral: códigos de negociação das companhias
CVM_FUNDOS = "cvm_fundos"  # informes mensais de FII e de Fiagro
BCB = "bcb"
B3 = "b3"
ORIGENS_CVM = frozenset({CVM_FRE, CVM_IPE})


def partes_da_origem(origem: str) -> frozenset[str]:
    """``"b3+cvm_fre"`` -> {"b3", "cvm_fre"}; vazio vale B3 (pacotes antigos, sem a coluna)."""
    return frozenset(origem.split("+")) if origem else frozenset({B3})


@dataclass(frozen=True, slots=True)
class Evento:
    """Evento corporativo (desdobramento, grupamento, bonificação, subscrição...)."""

    emissor: str
    isin: str
    tipo: str
    fator: str
    data_com: str
    aprovado_em: str
    isin_emitido: str
    preco: str
    periodo: str
    data_subscricao: str
    origem: str
    referencia: str
    custo_atribuido: str

    @classmethod
    def de_linha(cls, r: Mapping[str, str]) -> "Evento":
        return cls(**{c.name: r.get(c.name, "") for c in fields(cls)})

    @property
    def chave(self) -> tuple[str, str, str, str, str]:
        """Identidade do evento, na ordem em que a tabela vem ordenada."""
        return (self.data_com, self.isin, self.tipo, self.isin_emitido, self.fator)

    @property
    def multiplicador(self) -> Decimal:
        fator = Decimal(self.fator)
        return fator if self.tipo == "GRUPAMENTO" else 1 + fator / 100


def evento_valido(r: Mapping[str, str]) -> bool:
    """Linha de evento utilizável: ISIN, data com ISO e fator positivo."""
    try:
        date.fromisoformat(r["data_com"])
        fator = Decimal(r["fator"])
        _ = Decimal(r["preco"]) if r["preco"] else None
        if r["data_subscricao"]:
            date.fromisoformat(r["data_subscricao"])
    except (KeyError, ValueError, ArithmeticError):
        return False
    return bool(r["isin"]) and fator.is_finite() and fator > 0


@dataclass(frozen=True, slots=True)
class Provento:
    """Provento em dinheiro anunciado (com ISIN e, quando já definido, a data de pagamento)."""

    emissor: str
    isin: str
    tipo: str
    valor: str
    aprovado_em: str
    data_com: str
    pagamento: str
    referente: str
    origem: str = B3

    @classmethod
    def de_linha(cls, r: Mapping[str, str], origem: str = B3) -> "Provento":
        campos = {c.name: r.get(c.name, "") for c in fields(cls) if c.name != "origem"}
        return cls(**campos, origem=origem)


@dataclass(frozen=True, slots=True)
class ProventoHistorico:
    """JCP e restituição de capital de anos antigos: empresa e classe, sem data de pagamento."""

    emissor: str
    classe: str
    tipo: str
    valor: str
    aprovado_em: str
    data_com: str
    origem: str = B3

    @classmethod
    def de_linha(cls, r: Mapping[str, str], origem: str = B3) -> "ProventoHistorico":
        campos = {c.name: r.get(c.name, "") for c in fields(cls) if c.name != "origem"}
        return cls(**campos, origem=origem)


@dataclass(frozen=True, slots=True)
class Instrumento:
    """Ticker -> emissor, especificação ("ON NM") e razão social, e até quando foi visto."""

    ticker: str
    isin: str
    especificacao: str
    emissor: str
    nome: str
    visto_ate: str
    origem: str


@dataclass(frozen=True, slots=True)
class Emissor:
    """Emissor (4 letras) -> CNPJ, razão social e tipo (COMPANHIA, FII, ETF...)."""

    emissor: str
    cnpj: str
    razao_social: str
    tipo: str
    visto_ate: str
    origem: str


@dataclass(frozen=True, slots=True)
class NomePregao:
    """Nome de pregão de um ticker num período de negócios."""

    nome_pregao: str
    ticker: str
    desde: str
    ate: str
    origem: str


@dataclass(frozen=True, slots=True)
class TrocaCodigo:
    """Troca de código: o ticker impresso até ``ate`` e o de depois, com o ISIN de cada um."""

    ticker_impresso: str
    ticker: str
    ate: str
    isin_antigo: str
    isin_novo: str
    origem: str


@dataclass(frozen=True, slots=True)
class Divergencia:
    """Duas fontes com o mesmo registro e conteúdo diferente: vale o de maior precedência, e a
    diferença fica registrada (sem escolha calada)."""

    conceito: str
    chave: tuple[str, ...]
    mantida: str
    descartada: str
