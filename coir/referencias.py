"""Carrega as tabelas de referência de data/ (fora do parsing, que as recebe prontas)."""

import csv
from datetime import date
from decimal import Decimal
from functools import cache
from itertools import pairwise
from pathlib import Path

from coir.domain.fiscal import TipoAtivo
from coir.domain.parametros import (
    Calendario,
    IsencaoPME,
    Parametro,
    ParametrosFiscais,
    ReferenciasFiscais,
    RegraVencimento,
    TiposAtivo,
)
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
            moeda=c["moeda"],  # type: ignore[arg-type]  (validado pelo modelo)
            liquidacao=c["liquidacao"],  # type: ignore[arg-type]
            grupo=c["grupo"],
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
        instituicoes_b3=_mapa(base / "instituicoes_b3.csv", "prefixo", "corretora"),
    )


DIAS_SEMANA = {"segunda": 0, "terca": 1, "quarta": 2, "quinta": 3, "sexta": 4}


def _selic(arquivo: Path) -> dict[tuple[int, int], Decimal]:
    """Selic mensal (% a.m., BCB SGS 4390): meses completos, contíguos, sem repetição, 2 casas."""
    selic: dict[tuple[int, int], Decimal] = {}
    for x in _linhas(arquivo):
        try:
            mes = (int(x["mes"][:4]), int(x["mes"][5:7]))
            taxa = Decimal(x["taxa"])
        except (ValueError, ArithmeticError) as erro:
            raise ConfiguracaoError(f"{arquivo.name}: linha ilegível {x}") from erro
        if mes in selic or taxa < 0 or taxa != taxa.quantize(Decimal("0.01")):
            raise ConfiguracaoError(f"{arquivo.name}: {x['mes']} repetido ou taxa inválida")
        selic[mes] = taxa
    for (a1, m1), (a2, m2) in pairwise(sorted(selic)):
        if (a2, m2) != ((a1 + 1, 1) if m1 == 12 else (a1, m1 + 1)):
            raise ConfiguracaoError(f"{arquivo.name}: falta a Selic depois de {m1:02d}/{a1}")
    return selic


@cache
def carregar_fiscais(pasta: Path | None = None) -> ReferenciasFiscais:
    """Parâmetros fiscais, tipos de ativo e calendário (feriados e vencimento de opções)."""
    base = pasta or pasta_data()
    parametros = ParametrosFiscais(
        parametros=tuple(
            Parametro(nome=p["nome"], valor=p["valor"], inicio=_data(p["inicio"]),
                      fim=_data(p["fim"]), fonte=p["fonte"])
            for p in _linhas(base / "parametros_fiscais.csv")
        )
    )  # fmt: skip
    tipos: dict[str, TipoAtivo] = {}
    for t in _linhas(base / "tipos_ativo.csv"):
        if t["ticker"] in tipos:
            raise ConfiguracaoError(f"tipos_ativo.csv: {t['ticker']} repetido")
        tipos[t["ticker"]] = TipoAtivo(t["tipo"])
    feriados = _linhas(base / "feriados.csv")
    anos = {date.fromisoformat(f["data"]).year for f in feriados}
    regras = tuple(
        RegraVencimento(inicio=_data(r["inicio"]), fim=_data(r["fim"]),
                        dia_semana=DIAS_SEMANA[r["dia_semana"]], ordem=int(r["ordem"]),
                        ajuste=r["ajuste"], fonte=r["fonte"])
        for r in _linhas(base / "vencimentos_opcoes.csv")
    )  # fmt: skip
    excecoes = {
        e["chave"]: date.fromisoformat(e["data"])
        for e in _linhas(base / "vencimentos_opcoes_excecoes.csv")
    }
    calendario = Calendario(
        primeiro_ano=min(anos), ultimo_ano=max(anos),
        nao_uteis=frozenset(date.fromisoformat(f["data"]) for f in feriados
                            if f["tipo"] == "bancario"),
        sem_pregao=frozenset(date.fromisoformat(f["data"]) for f in feriados),
        vencimentos=regras, excecoes_vencimento=excecoes,
    )  # fmt: skip
    pme = {
        p["ticker"]: IsencaoPME(
            ticker=p["ticker"], aquisicao_desde=date.fromisoformat(p["aquisicao_desde"]),
            aquisicao_ate=date.fromisoformat(p["aquisicao_ate"]),
            vendas_ate=date.fromisoformat(p["vendas_ate"]), fonte=p["fonte"],
        )
        for p in _linhas(base / "isencao_pme.csv")
    }  # fmt: skip
    selic = _selic(base / "selic_mensal.csv")
    return ReferenciasFiscais(parametros=parametros, tipos=TiposAtivo(tabela=tipos),
                              calendario=calendario, pme=pme, selic=selic)  # fmt: skip
