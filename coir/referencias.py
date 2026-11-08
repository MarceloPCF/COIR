"""Carrega as tabelas de referência de data/ (fora do parsing, que as recebe prontas).

As tabelas geradas pelo pipeline público (``*_b3.csv``, F7b) entram por baixo das curadas à mão:
nome de pregão, raiz de opção ou tipo de ativo da tabela manual sempre vale. A pasta pode ser a
que veio na instalação ou o pacote mais novo baixado em ``<dados>/referencias/atual``
(``pasta_efetiva``).
"""

import csv
import json
import re
from datetime import date
from decimal import Decimal
from functools import cache
from itertools import pairwise
from pathlib import Path

import yaml
from pydantic import ValidationError

from coir.domain.cpf import cnpj_valido
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


def _mapa_gerado(
    caminho: Path, chave: str, valor: str, *, compacta: bool = False
) -> dict[str, str]:
    return _mapa(caminho, chave, valor, compacta=compacta) if caminho.is_file() else {}


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


RE_TICKER_PADRAO = re.compile(r"^[A-Z0-9]{4}([3-8]|11|5[3-6])$")  # 53 a 56: ações "P" (RAIZ55)


COLUNAS_EVENTOS_B3 = ("emissor", "isin", "tipo", "fator", "data_com", "aprovado_em",
                      "isin_emitido", "preco", "periodo", "data_subscricao")  # fmt: skip
COLUNAS_INSTRUMENTOS_B3 = ("ticker", "isin", "categoria", "especificacao", "emissor", "nome",
                           "visto_desde", "visto_ate")  # fmt: skip
COLUNAS_EMISSORES_B3 = ("emissor", "cnpj", "razao_social", "tipo", "visto_desde", "visto_ate")
COLUNAS_EMISSORES = ("emissor", "cnpj", "razao_social", "observacao")
COLUNAS_PROVENTOS_B3 = ("emissor", "isin", "tipo", "valor", "aprovado_em", "data_com", "pagamento",
                        "referente")  # fmt: skip
COLUNAS_HISTORICO_B3 = ("emissor", "classe", "tipo", "valor", "aprovado_em", "data_com")


def evento_b3_valido(r: dict[str, str]) -> bool:
    """Linha de ``eventos_b3.csv`` utilizável: ISIN, data com ISO e fator positivo."""
    try:
        date.fromisoformat(r["data_com"])
        fator = Decimal(r["fator"])
        _ = Decimal(r["preco"]) if r["preco"] else None
        if r["data_subscricao"]:
            date.fromisoformat(r["data_subscricao"])
    except (KeyError, ValueError, ArithmeticError):
        return False
    return bool(r["isin"]) and fator.is_finite() and fator > 0


def eventos_b3(pasta: Path) -> list[dict[str, str]]:
    """Eventos corporativos do pacote público (``eventos_b3.csv``; vazio se não houver), só as
    linhas utilizáveis e sem os pares que se anulam (``_neutros``)."""
    caminho = pasta / "eventos_b3.csv"
    validas = [r for r in (_linhas(caminho) if caminho.is_file() else []) if evento_b3_valido(r)]
    neutros = _neutros(validas)
    return [r for i, r in enumerate(validas) if i not in neutros]


def _multiplicador(r: dict[str, str]) -> Decimal:
    fator = Decimal(r["fator"])
    return fator if r["tipo"] == "GRUPAMENTO" else 1 + fator / 100


def _neutros(linhas: list[dict[str, str]]) -> set[int]:
    """Grupamento e desdobramento no mesmo ISIN e data com que se anulam (ITUB 2011: 100 para 1 e
    1 para 100): a quantidade não muda, então não entram em controle nenhum. Ficam no pacote (que
    nunca apaga): se a B3 voltar a listar só um dos lados, a coleta não o toma por evento novo."""
    grupos: dict[tuple[str, str, str], list[int]] = {}
    for i, r in enumerate(linhas):
        if r["tipo"] in ("GRUPAMENTO", "DESDOBRAMENTO", "BONIFICACAO"):
            grupos.setdefault((r["isin"], r["data_com"], r["isin_emitido"]), []).append(i)
    neutros: set[int] = set()
    for indices in grupos.values():
        grupamento = {linhas[i]["tipo"] == "GRUPAMENTO" for i in indices}
        produto = Decimal(1)
        for i in indices:
            produto *= _multiplicador(linhas[i])
        if grupamento == {True, False} and produto == 1:
            neutros.update(indices)
    return neutros


def validar_tabelas_b3(pasta: Path) -> None:
    """Pacote novo: as tabelas que só as propostas da tela 7 leem também têm de estar íntegras."""
    for nome, colunas in (("eventos_b3.csv", COLUNAS_EVENTOS_B3),
                          ("instrumentos_b3.csv", COLUNAS_INSTRUMENTOS_B3),
                          ("emissores_b3.csv", COLUNAS_EMISSORES_B3),
                          ("proventos_b3.csv", COLUNAS_PROVENTOS_B3),
                          ("proventos_historico_b3.csv", COLUNAS_HISTORICO_B3)):  # fmt: skip
        caminho = pasta / nome
        if not caminho.is_file():
            continue
        linhas = _linhas(caminho)
        if linhas and set(colunas) - set(linhas[0]):
            raise ConfiguracaoError(f"{nome}: colunas ausentes")
        if nome == "eventos_b3.csv":
            ruins = [r for r in linhas if not evento_b3_valido(r)]
            if ruins:
                raise ConfiguracaoError(f"{nome}: {len(ruins)} linha(s) inválida(s)")
        elif nome == "emissores_b3.csv":
            if any(not r["emissor"] or not cnpj_valido(r["cnpj"]) for r in linhas):
                raise ConfiguracaoError(f"{nome}: linha sem emissor ou com CNPJ inválido")
        elif nome == "proventos_b3.csv":
            if len(proventos_b3(pasta)) != len(linhas):
                raise ConfiguracaoError(f"{nome}: linha com data ou valor inválido")
        elif nome == "proventos_historico_b3.csv":
            if len(proventos_historico_b3(pasta)) != len(linhas):
                raise ConfiguracaoError(f"{nome}: linha com data ou valor inválido")
        elif any(not r["ticker"] or not r["isin"] for r in linhas):
            raise ConfiguracaoError(f"{nome}: linha sem ticker ou ISIN")


def tickers_por_isin(pasta: Path) -> dict[str, str]:
    """ISIN -> ticker de ação ou unit (o mais recente visto no cadastro de instrumentos). O ISIN de
    antes de uma troca de código (``trocas_ticker.csv``: KROT3 era BRKROTACNOR9) leva ao ticker
    da época, que o cadastro de hoje não tem: o evento daquele ISIN acha a posição do ativo."""
    caminho = pasta / "instrumentos_b3.csv"
    saida: dict[str, tuple[str, str]] = {}
    for r in _linhas(caminho) if caminho.is_file() else []:
        if RE_TICKER_PADRAO.match(r["ticker"]) and r["visto_ate"] >= saida.get(
            r["isin"], ("", ""))[0]:  # fmt: skip
            saida[r["isin"]] = (r["visto_ate"], r["ticker"])
    por_isin = {isin: t for isin, (_, t) in saida.items()}
    trocas = pasta / "trocas_ticker.csv"
    for t in _linhas(trocas) if trocas.is_file() else []:
        for isin, ticker in ((t.get("isin_antigo", ""), t["ticker_impresso"]),
                             (t.get("isin_novo", ""), t["ticker"])):  # fmt: skip
            if isin and RE_TICKER_PADRAO.match(ticker):
                por_isin.setdefault(isin, ticker)
    return por_isin


def emissores(pasta: Path) -> dict[str, tuple[str, str, str]]:
    """Emissor (4 letras) -> (CNPJ, razão social, tipo da B3: COMPANHIA, FII, FIAGRO, ETF...),
    para o Bens e Direitos. O pacote público (``emissores_b3.csv``, o CNPJ visto por último) fica
    por baixo da tabela curada ``emissores.csv``, que sempre vale (o tipo continua o do pacote).
    CNPJ com dígito errado na tabela curada falha alto."""
    saida: dict[str, tuple[str, str, str, str]] = {}
    gerado = pasta / "emissores_b3.csv"
    for r in _linhas(gerado) if gerado.is_file() else []:
        visto = saida.get(r["emissor"], ("", "", "", ""))[3]
        if cnpj_valido(r["cnpj"]) and r["visto_ate"] >= visto:
            saida[r["emissor"]] = (r["cnpj"], r["razao_social"], r["tipo"], r["visto_ate"])
    curado = pasta / "emissores.csv"
    for r in _linhas(curado) if curado.is_file() else []:
        if not cnpj_valido(r["cnpj"]):
            raise ConfiguracaoError(f"emissores.csv: CNPJ inválido para {r['emissor']}")
        tipo = saida.get(r["emissor"], ("", "", "", ""))[2]
        saida[r["emissor"]] = (r["cnpj"], r["razao_social"], tipo, "")
    return {e: (c, n, t) for e, (c, n, t, _) in saida.items()}


def proventos_b3(pasta: Path) -> list[dict[str, str]]:
    """Proventos em dinheiro do pacote público (``proventos_b3.csv``; vazio se não houver), só as
    linhas com data de aprovação válida, pagamento válido ou ainda a definir (vazio) e valor
    numérico."""
    caminho = pasta / "proventos_b3.csv"
    saida: list[dict[str, str]] = []
    for r in _linhas(caminho) if caminho.is_file() else []:
        try:
            date.fromisoformat(r["aprovado_em"])
            if r["pagamento"]:
                date.fromisoformat(r["pagamento"])
            if not Decimal(r["valor"]).is_finite():
                continue
        except (KeyError, ValueError, ArithmeticError):
            continue
        saida.append(r)
    return saida


def proventos_historico_b3(pasta: Path) -> list[dict[str, str]]:
    """JCP do histórico público (``proventos_historico_b3.csv``: classe, aprovação e data com, sem
    pagamento; vazio se não houver), só as linhas com datas válidas e valor numérico."""
    caminho = pasta / "proventos_historico_b3.csv"
    saida: list[dict[str, str]] = []
    for r in _linhas(caminho) if caminho.is_file() else []:
        try:
            date.fromisoformat(r["aprovado_em"])
            date.fromisoformat(r["data_com"])
            if not r["classe"] or not Decimal(r["valor"]).is_finite():
                continue
        except (KeyError, ValueError, ArithmeticError):
            continue
        saida.append(r)
    return saida


def instrumentos(pasta: Path) -> dict[str, tuple[str, str, str]]:
    """Ticker -> (emissor, especificação "ON NM", razão social), do cadastro de instrumentos da
    B3 (``instrumentos_b3.csv``; vazio se não houver), o visto por último."""
    caminho = pasta / "instrumentos_b3.csv"
    saida: dict[str, tuple[str, str, str, str]] = {}
    for r in _linhas(caminho) if caminho.is_file() else []:
        if r["visto_ate"] >= saida.get(r["ticker"], ("", "", "", ""))[3]:
            saida[r["ticker"]] = (r["emissor"], r["especificacao"], r["nome"], r["visto_ate"])
    return {t: (e, s, n) for t, (e, s, n, _) in saida.items()}


@cache
def ultimos_negocios(pasta: Path) -> dict[str, date]:
    """Ticker -> último negócio no COTAHIST (``nomes_pregao_b3.csv``; vazio em pacote antigo)."""
    caminho = pasta / "nomes_pregao_b3.csv"
    saida: dict[str, date] = {}
    for r in _linhas(caminho) if caminho.is_file() else []:
        ate = date.fromisoformat(r["ate"])
        saida[r["ticker"]] = max(ate, saida.get(r["ticker"], ate))
    return saida


COLUNAS_SAIDAS = ("ticker", "ultimo_negocio", "situacao", "destino", "relacao",
                  "dinheiro_por_acao", "confianca", "fonte", "observacao")  # fmt: skip


@cache
def saidas_b3(pasta: Path) -> dict[str, dict[str, str]]:
    """O que houve com o ativo que parou de negociar (``saidas_b3.csv``, curada, com fonte)."""
    caminho = pasta / "saidas_b3.csv"
    linhas = _linhas(caminho) if caminho.is_file() else []
    if linhas and set(COLUNAS_SAIDAS) - set(linhas[0]):
        raise ConfiguracaoError("saidas_b3.csv: colunas ausentes")
    return {r["ticker"]: r for r in linhas}


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
        mapa = MapeamentoIrpf.model_validate(bruto)
    except (yaml.YAMLError, ValidationError) as erro:
        raise ConfiguracaoError(f"data/irpf/{exercicio}.yaml inválido: {erro}") from erro
    if mapa.exercicio != exercicio:
        raise ConfiguracaoError(f"data/irpf/{exercicio}.yaml traz o exercício {mapa.exercicio}")
    return mapa


def _data(texto: str) -> date | None:
    return date.fromisoformat(texto) if texto else None


def _nomes_historicos(caminho: Path) -> dict[str, tuple[tuple[str, str, str], ...]]:
    """Histórico de nomes de pregão da B3 (``nomes_pregao_b3.csv``; vazio em pacote antigo)."""
    saida: dict[str, list[tuple[str, str, str]]] = {}
    for r in _linhas(caminho) if caminho.is_file() else []:
        saida.setdefault(compactar(r["nome_pregao"]), []).append(
            (r["desde"], r["ate"], r["ticker"])
        )
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
        for c in _linhas(base / "contratos_futuros.csv")
    )  # fmt: skip
    return Referencias(
        ativos=_mapa_gerado(base / "ativos_b3.csv", "nome_pregao", "ticker", compacta=True)
        | _mapa(base / "ativos.csv", "nome_pregao", "ticker", compacta=True),
        opcoes_raiz=_mapa_gerado(base / "opcoes_raiz_b3.csv", "raiz", "ticker")
        | _mapa(base / "opcoes_raiz.csv", "raiz", "ticker"),
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
        nomes_historicos=_nomes_historicos(base / "nomes_pregao_b3.csv"),
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
    gerados = base / "tipos_ativo_b3.csv"  # por baixo da tabela manual
    for t in _linhas(gerados) if gerados.is_file() else []:
        tipos.setdefault(t["ticker"], TipoAtivo(t["tipo"]))
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
            media_fim=Decimal(p["media_fim"]) if p.get("media_fim") else None,
        )
        for p in _linhas(base / "isencao_pme.csv")
    }  # fmt: skip
    selic = _selic(base / "selic_mensal.csv")
    return ReferenciasFiscais(parametros=parametros, tipos=TiposAtivo(tabela=tipos),
                              calendario=calendario, pme=pme, selic=selic)  # fmt: skip
