"""Parâmetros fiscais com vigência e tipos de ativo (de data/, carregados fora do domínio).

Regra 2 do CLAUDE.md: nenhum parâmetro fiscal no código. O motor recebe estes objetos prontos.
"""

import re
from datetime import date
from decimal import Decimal

from pydantic import BaseModel, ConfigDict

from coir.domain.fiscal import TipoAtivo
from coir.erros import ReferenciaError

RE_ACAO = re.compile(r"^[A-Z0-9]{4}[3-8]$")
RE_BDR = re.compile(r"^[A-Z0-9]{4}3[1-9]$")


class Modelo(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")


class Parametro(Modelo):
    nome: str
    valor: str
    inicio: date | None = None
    fim: date | None = None
    fonte: str = ""

    def vigente(self, dia: date) -> bool:
        return (self.inicio is None or self.inicio <= dia) and (self.fim is None or dia <= self.fim)


class ParametrosFiscais(Modelo):
    parametros: tuple[Parametro, ...]

    def _um(self, nome: str, dia: date) -> Parametro:
        achados = [p for p in self.parametros if p.nome == nome and p.vigente(dia)]
        if len(achados) != 1:
            situacao = "sem vigência" if not achados else "com vigências sobrepostas"
            raise ReferenciaError(f"parâmetro fiscal {nome!r} {situacao} em {dia.isoformat()}")
        return achados[0]

    def decimal(self, nome: str, dia: date) -> Decimal:
        return Decimal(self._um(nome, dia).valor)

    def booleano(self, nome: str, dia: date) -> bool:
        valor = self._um(nome, dia).valor.lower()
        if valor not in ("true", "false"):
            raise ReferenciaError(f"parâmetro fiscal {nome!r}: esperado true/false, lido {valor!r}")
        return valor == "true"

    def texto(self, nome: str, dia: date) -> str:
        return self._um(nome, dia).valor


class TiposAtivo(Modelo):
    """Ticker -> tipo. Sufixo 3 a 8: ação; 31 a 39: BDR; final 11 (unit, ETF, FII) e exceções
    vêm da tabela (sem palpite: ticker sem tipo é erro)."""

    tabela: dict[str, TipoAtivo]

    def tipo(self, ticker: str) -> TipoAtivo:
        if ticker in self.tabela:
            return self.tabela[ticker]
        if RE_ACAO.match(ticker):
            return TipoAtivo.ACAO
        if RE_BDR.match(ticker):
            return TipoAtivo.BDR
        raise ReferenciaError(
            f"ticker {ticker} sem tipo de ativo: cadastre em data/tipos_ativo.csv "
            "(UNIT, ETF, ETF_RF, BDR, FII ou ACAO)"
        )


class RegraVencimento(Modelo):
    """Vencimento de opções sobre ações: n-ésimo dia da semana do mês; sem pregão nesse dia,
    o pregão anterior ou o seguinte."""

    inicio: date | None = None
    fim: date | None = None
    dia_semana: int  # 0 = segunda ... 4 = sexta (date.weekday)
    ordem: int  # 3 = terceira
    ajuste: str  # "anterior" ou "seguinte"
    fonte: str = ""

    def vigente(self, dia: date) -> bool:
        return (self.inicio is None or self.inicio <= dia) and (self.fim is None or dia <= self.fim)


class Calendario(Modelo):
    """Dias sem expediente bancário (vencimento de DARF) e sem pregão (opções, termo), para os
    anos de `primeiro_ano` a `ultimo_ano` (fora deles é erro, nunca "sem feriado")."""

    primeiro_ano: int
    ultimo_ano: int
    nao_uteis: frozenset[date]  # feriados bancários (fins de semana à parte)
    sem_pregao: frozenset[date]  # dias úteis ou feriados sem pregão na B3 (fins de semana à parte)
    vencimentos: tuple[RegraVencimento, ...] = ()
    excecoes_vencimento: dict[str, date] = {}  # "AAAA-MM" -> data remarcada pela B3

    def cobre(self, dia: date) -> None:
        if not self.primeiro_ano <= dia.year <= self.ultimo_ano:
            raise ReferenciaError(
                f"calendário de feriados sem o ano {dia.year} (data/feriados.csv cobre "
                f"{self.primeiro_ano}-{self.ultimo_ano})"
            )


class IsencaoPME(Modelo):
    """Ações de pequenas e médias empresas (Lei 13.043/2014 art. 16; IN RFB 1.585/2015 arts. 66 a
    69-A): ganho isento nas vendas até `vendas_ate` das ações compradas entre `aquisicao_desde` e
    `aquisicao_ate` (antes da oferta subsequente que desenquadrou a companhia)."""

    ticker: str
    aquisicao_desde: date
    aquisicao_ate: date
    vendas_ate: date
    fonte: str = ""

    def compra_elegivel(self, dia: date) -> bool:
        return self.aquisicao_desde <= dia <= self.aquisicao_ate


class ReferenciasFiscais(Modelo):
    parametros: ParametrosFiscais
    tipos: TiposAtivo
    calendario: Calendario
    pme: dict[str, IsencaoPME] = {}
    # Selic acumulada no mês (% a.m.), por (ano, mês): juros de mora da DARF (Lei 9.430/1996
    # art. 61 §3º e art. 5º §3º). data/selic_mensal.csv, só meses completos (BCB, SGS 4390).
    selic: dict[tuple[int, int], Decimal] = {}
