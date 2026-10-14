"""Tabelas de referência usadas na normalização (carregadas de data/ fora do domínio)."""

from datetime import date
from decimal import Decimal

from pydantic import BaseModel, ConfigDict


class ContratoFuturo(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")

    raiz: str
    mercado: str
    multiplicador: Decimal
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

    def ticker_na_data(self, ticker: str, dia: date) -> str:
        for t in self.trocas:
            if t.ticker_impresso == ticker and dia <= t.ate:
                return t.ticker
        return ticker
