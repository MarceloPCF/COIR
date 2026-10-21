"""Tabelas de referência usadas na normalização (carregadas de data/ fora do domínio)."""

import re
import unicodedata
from collections.abc import Callable
from datetime import date
from decimal import Decimal
from typing import Literal

from pydantic import BaseModel, ConfigDict

RE_SERIE_OPCAO = re.compile(r"^[A-Z0-9]{4}[A-X]\d")


def opcao_de_compra(serie: str) -> bool:
    """Letra do mês da série: A a L, opção de compra (call); M a X, de venda (put)."""
    return serie[4:5] <= "L"


def especie_opcao(serie: str) -> str:
    return "compra" if opcao_de_compra(serie) else "venda"


class ContratoFuturo(BaseModel):
    """Contrato futuro da B3 (data/contratos_futuros.csv).

    ``moeda``: da cotação (café e soja CME em US$; o ajuste impresso na nota já vem em R$).
    ``liquidacao``: FISICA = entrega da mercadoria no vencimento (fora do escopo: o COIR avisa).
    ``grupo``: linha da apuração e do IRPF (indice, dolar, commodity = "outros")."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    raiz: str
    mercado: str
    multiplicador: Decimal
    moeda: Literal["BRL", "USD"] = "BRL"
    liquidacao: Literal["FINANCEIRA", "FISICA"] = "FINANCEIRA"
    grupo: str = "indice"
    inicio: date | None = None
    fim: date | None = None

    def vigente(self, dia: date) -> bool:
        return (self.inicio is None or self.inicio <= dia) and (self.fim is None or dia <= self.fim)


class TrocaTicker(BaseModel):
    """O ticker impresso vale como outro até uma data (ex.: JSLG3 era SIMH3 até 18/09/2020)."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    ticker_impresso: str
    ticker: str
    ate: date


class Referencias(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")

    ativos: dict[str, str]  # nome de pregão compacto (sem espaços, minúsculo) -> ticker
    opcoes_raiz: dict[str, str]  # raiz de 4 letras -> ticker do ativo-objeto (último recurso)
    classes: dict[str, str]  # ON -> 3, PN -> 4, ..., UNT -> 11
    mercados_bovespa: dict[str, str]  # tipo de mercado impresso -> mercado normalizado
    contratos: tuple[ContratoFuturo, ...]
    sufixos: tuple[str, ...]  # sufixos de governança removidos do nome de pregão (compactos)
    trocas: tuple[TrocaTicker, ...] = ()
    instituicoes_b3: dict[str, str] = {}  # prefixo do nome compactado -> corretora (perfil)

    def corretora_b3(self, instituicao: str) -> str | None:
        """Instituição como a B3 escreve ("XP INVESTIMENTOS CCTVM S/A.") -> corretora do COIR."""
        chave = re.sub(r"[^A-Z0-9]", "", unicodedata.normalize("NFKD", instituicao.upper())
                       .encode("ascii", "ignore").decode())  # fmt: skip
        return next((c for p, c in self.instituicoes_b3.items() if chave.startswith(p)), None)

    def ticker_na_data(self, ticker: str, dia: date) -> str:
        for t in self.trocas:
            if t.ticker_impresso == ticker and dia <= t.ate:
                return t.ticker
        if RE_SERIE_OPCAO.match(ticker):  # a série segue a troca do ativo-objeto
            for t in self.trocas:
                if ticker[:4] == t.ticker_impresso[:4] and dia <= t.ate:
                    return t.ticker[:4] + ticker[4:]  # RRRPT32 -> BRAVT32
        return ticker


def codigo_atual(refs: Referencias) -> Callable[[str], str]:
    """O código mais recente do ativo pela tabela de trocas (planilhas e declarações usam ora o
    antigo, ora o novo; a série de opção segue a raiz)."""

    def atual(ticker: str) -> str:
        for _ in range(len(refs.trocas) + 1):
            novo = refs.ticker_na_data(ticker, date.min)
            if novo == ticker:
                break
            ticker = novo
        return ticker

    return atual
