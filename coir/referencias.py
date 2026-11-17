"""Carrega as tabelas de referência de data/ que o motor e o parsing recebem prontas.

As tabelas da bolsa (provedor ``coir.fontes.provedores.b3``) entram por baixo das curadas à mão:
nome de pregão, raiz de opção ou tipo de ativo da tabela manual sempre vale. Eventos, proventos e
cadastros que só os serviços leem ficam em ``coir.fontes``, por conceito. A pasta pode ser a
que veio na instalação ou o pacote mais novo baixado em ``<dados>/referencias/atual``
(``pasta_efetiva``).
"""

import json
from datetime import date
from decimal import Decimal
from functools import cache
from pathlib import Path

import yaml
from pydantic import ValidationError

from coir.domain.fiscal import TipoAtivo
from coir.domain.irpf import MapeamentoIrpf
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
from coir.fontes.leitura import linhas, mapa
from coir.fontes.provedores import b3, bcb
from coir.parsing.numeros import compactar
from coir.recursos import pasta_data

FORMATO = 2  # esquema de data/ que esta versão lê (igual ao de tools/referencias/gerar.py)
PACOTE = "pacote.json"
PASTA_BAIXADA = Path("referencias") / "atual"


def info_pacote(pasta: Path) -> tuple[int, date] | None:
    """(formato, data) de ``pacote.json``; None se a pasta não tem pacote legível."""
    try:
        info = json.loads((pasta / PACOTE).read_text(encoding="utf-8"))
        return int(info["formato"]), date.fromisoformat(str(info["data"]))
    except (OSError, ValueError, KeyError, TypeError):
        return None


def pasta_efetiva(dados: Path | None) -> Path:
    """O pacote baixado (``<dados>/referencias/atual``) se for do formato desta versão e mais novo
    que o da instalação; senão, o ``data/`` da instalação (sem rede, vale o que veio nela)."""
    embutida = pasta_data()
    if dados is None:
        return embutida
    baixada = dados / PASTA_BAIXADA
    b, e = info_pacote(baixada), info_pacote(embutida)
    if b is not None and b[0] == FORMATO and (e is None or b[1] > e[1]):
        return baixada
    return embutida


@cache
def carregar_irpf(ano_calendario: int, pasta: Path | None = None) -> MapeamentoIrpf:
    """Mapeamento das fichas do programa IRPF (``data/irpf/<exercício>.yaml``). Vem sempre com a
    versão instalada (não com o pacote público): cada exercício novo pede código novo no COIR.
    Exercício sem arquivo falha alto."""
    exercicio = ano_calendario + 1
    caminho = (pasta or pasta_data()) / "irpf" / f"{exercicio}.yaml"
    if not caminho.is_file():
        raise ConfiguracaoError(f"sem mapeamento das fichas do IRPF {exercicio} (ano-calendário "
                                f"{ano_calendario}): falta data/irpf/{exercicio}.yaml")  # fmt: skip
    try:
        bruto = yaml.safe_load(caminho.read_text(encoding="utf-8"))
        mapeamento = MapeamentoIrpf.model_validate(bruto)
    except (yaml.YAMLError, ValidationError) as erro:
        raise ConfiguracaoError(f"data/irpf/{exercicio}.yaml inválido: {erro}") from erro
    if mapeamento.exercicio != exercicio:
        raise ConfiguracaoError(
            f"data/irpf/{exercicio}.yaml traz o exercício {mapeamento.exercicio}"
        )
    return mapeamento


def _data(texto: str) -> date | None:
    return date.fromisoformat(texto) if texto else None


def _nomes_historicos(pasta: Path) -> dict[str, tuple[tuple[str, str, str], ...]]:
    """Histórico de nomes de pregão (vazio em pacote antigo ou sem a fonte)."""
    saida: dict[str, list[tuple[str, str, str]]] = {}
    for n in b3.nomes_pregao(pasta):
        saida.setdefault(compactar(n.nome_pregao), []).append((n.desde, n.ate, n.ticker))
    return {nome: tuple(periodos) for nome, periodos in saida.items()}


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
        for c in linhas(base / "contratos_futuros.csv")
    )  # fmt: skip
    return Referencias(
        ativos=b3.ativos(base) | mapa(base / "ativos.csv", "nome_pregao", "ticker", compacta=True),
        opcoes_raiz=b3.opcoes_raiz(base) | mapa(base / "opcoes_raiz.csv", "raiz", "ticker"),
        classes=mapa(base / "classes_acao.csv", "classe", "digito"),
        mercados_bovespa=mapa(base / "mercados_bovespa.csv", "tipo_mercado", "mercado"),
        contratos=contratos,
        sufixos=tuple(compactar(s["sufixo"]) for s in linhas(base / "sufixos_especificacao.csv")),
        trocas=tuple(
            TrocaTicker(
                ticker_impresso=t["ticker_impresso"],
                ticker=t["ticker"],
                ate=date.fromisoformat(t["ate"]),
            )
            for t in linhas(base / "trocas_ticker.csv")
        ),
        instituicoes_b3=mapa(base / "instituicoes_b3.csv", "prefixo", "corretora"),
        nomes_historicos=_nomes_historicos(base),
    )


DIAS_SEMANA = {"segunda": 0, "terca": 1, "quarta": 2, "quinta": 3, "sexta": 4}


@cache
def carregar_fiscais(pasta: Path | None = None) -> ReferenciasFiscais:
    """Parâmetros fiscais, tipos de ativo e calendário (feriados e vencimento de opções)."""
    base = pasta or pasta_data()
    parametros = ParametrosFiscais(
        parametros=tuple(
            Parametro(nome=p["nome"], valor=p["valor"], inicio=_data(p["inicio"]),
                      fim=_data(p["fim"]), fonte=p["fonte"])
            for p in linhas(base / "parametros_fiscais.csv")
        )
    )  # fmt: skip
    tipos: dict[str, TipoAtivo] = {}
    for t in linhas(base / "tipos_ativo.csv"):
        if t["ticker"] in tipos:
            raise ConfiguracaoError(f"tipos_ativo.csv: {t['ticker']} repetido")
        tipos[t["ticker"]] = TipoAtivo(t["tipo"])
    for ticker, tipo in b3.tipos_ativo(base):  # por baixo da tabela manual
        tipos.setdefault(ticker, TipoAtivo(tipo))
    feriados = linhas(base / "feriados.csv")
    anos = {date.fromisoformat(f["data"]).year for f in feriados}
    regras = tuple(
        RegraVencimento(inicio=_data(r["inicio"]), fim=_data(r["fim"]),
                        dia_semana=DIAS_SEMANA[r["dia_semana"]], ordem=int(r["ordem"]),
                        ajuste=r["ajuste"], fonte=r["fonte"])
        for r in linhas(base / "vencimentos_opcoes.csv")
    )  # fmt: skip
    excecoes = {
        e["chave"]: date.fromisoformat(e["data"])
        for e in linhas(base / "vencimentos_opcoes_excecoes.csv")
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
            media_fim=Decimal(p["media_fim"]) if p.get("media_fim") else None,
        )
        for p in linhas(base / "isencao_pme.csv")
    }  # fmt: skip
    selic = bcb.selic(base)
    return ReferenciasFiscais(parametros=parametros, tipos=TiposAtivo(tabela=tipos),
                              calendario=calendario, pme=pme, selic=selic)  # fmt: skip
