"""Carrega as tabelas de referência de data/ (fora do parsing, que as recebe prontas)."""

import csv
from datetime import date
from decimal import Decimal
from functools import cache
from pathlib import Path

from coir.domain.referencias import ContratoFuturo, Referencias, TrocaTicker
from coir.erros import ConfiguracaoError
from coir.parsing.numeros import compactar
from coir.recursos import pasta_data


def _linhas(caminho: Path) -> list[dict[str, str]]:
    if not caminho.is_file():
        raise ConfiguracaoError(f"tabela de referência ausente: {caminho.name}")
    with caminho.open(encoding="utf-8", newline="") as arquivo:
        return [
            {k: (v or "").strip() for k, v in linha.items()} for linha in csv.DictReader(arquivo)
        ]


def _mapa(caminho: Path, chave: str, valor: str, *, compacta: bool = False) -> dict[str, str]:
    mapa: dict[str, str] = {}
    for linha in _linhas(caminho):
        k = compactar(linha[chave]) if compacta else linha[chave]
        if k in mapa and mapa[k] != linha[valor]:
            raise ConfiguracaoError(f"{caminho.name}: {linha[chave]!r} com dois valores")
        mapa[k] = linha[valor]
    return mapa


def _data(texto: str) -> date | None:
    return date.fromisoformat(texto) if texto else None


@cache
def carregar(pasta: Path | None = None) -> Referencias:
    base = pasta or pasta_data()
    contratos = tuple(
        ContratoFuturo(
            raiz=c["raiz"],
            mercado=c["mercado"],
            multiplicador=Decimal(c["multiplicador"]),
            inicio=_data(c["inicio"]),
            fim=_data(c["fim"]),
        )
        for c in _linhas(base / "contratos_futuros.csv")
    )  # fmt: skip
    return Referencias(
        ativos=_mapa(base / "ativos.csv", "nome_pregao", "ticker", compacta=True),
        opcoes_raiz=_mapa(base / "opcoes_raiz.csv", "raiz", "ticker"),
        classes=_mapa(base / "classes_acao.csv", "classe", "digito"),
        mercados_bovespa=_mapa(base / "mercados_bovespa.csv", "tipo_mercado", "mercado"),
        contratos=contratos,
        sufixos=tuple(compactar(s["sufixo"]) for s in _linhas(base / "sufixos_especificacao.csv")),
        trocas=tuple(
            TrocaTicker(
                ticker_impresso=t["ticker_impresso"],
                ticker=t["ticker"],
                ate=date.fromisoformat(t["ate"]),
            )
            for t in _linhas(base / "trocas_ticker.csv")
        ),
    )
