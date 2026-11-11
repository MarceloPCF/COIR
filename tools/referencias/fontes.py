"""Download das fontes públicas (B3 e Banco Central), só com a biblioteca padrão.

Nenhum dado do usuário sai daqui: são consultas GET a páginas públicas, sem login. Respostas fora
do formato falham alto (o pacote não é gerado com dado pela metade).
"""

from __future__ import annotations

import base64
import json
import logging
import time
import urllib.error
import urllib.request
import zipfile
from collections.abc import Callable, Collection, Iterable, Mapping
from datetime import date, timedelta
from pathlib import Path
from typing import Any, cast

from tools.referencias import historico as hist
from tools.referencias.gerar import (
    Coleta,
    EventoB3,
    GeracaoError,
    ler_csv,
    ler_instrumentos,
    nomes_de_pregao,
)
from tools.referencias.historico import Revisao

log = logging.getLogger(__name__)
AGENTE = "COIR/3.0 (tabelas de referencia publicas; github.com/MarceloPCF/COIR)"
B3_ARQUIVOS = "https://arquivos.b3.com.br/api/download"
B3_EMPRESAS = "https://sistemaswebb3-listados.b3.com.br/listedCompaniesProxy/CompanyCall"
B3_FUNDOS = "https://sistemaswebb3-listados.b3.com.br/fundsListedProxy/Search"
TIPOS_FUNDO = ("FII", "FIAGRO", "ETF", "ETF-RF", "ETF-FII", "FI-INFRA", "FIP")  # viram ativo
# do FIP, só o de infraestrutura tem alíquota zero (Lei 11.478/2007 art. 2º); o comum paga 15%
FIP_ALIQUOTA_ZERO = {"FIP IE": "FIP-IE", "FIP PD&I": "FIP-PDI", "FIP PDI": "FIP-PDI"}
BCB_SELIC = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.4390/dados?formato=json"
CVM_FRE = "https://dados.cvm.gov.br/dados/CIA_ABERTA/DOC/FRE/DADOS/fre_cia_aberta_{ano}.zip"
B3_COTAHIST = "https://bvmf.bmfbovespa.com.br/InstDados/SerHist/COTAHIST_A{ano}.ZIP"
ANOS_FRE = range(2010, 2024)  # o FRE de 2024 em diante não tem mais a tabela de desdobramentos
CVM_FCA = "https://dados.cvm.gov.br/dados/CIA_ABERTA/DOC/FCA/DADOS/fca_cia_aberta_{ano}.zip"
FCA_DESDE = 2018  # o FCA traz o código de negociação de 2018 em diante
NOMES_DESDE = 2008  # primeiro ano com notas
Baixar = Callable[[str], bytes]


def baixar(url: str, tentativas: int = 3, pausa: float = 2.0) -> bytes:
    """GET com timeout e novas tentativas curtas (a B3 às vezes devolve 5xx)."""
    erro: Exception | None = None
    for n in range(tentativas):
        try:
            pedido = urllib.request.Request(url, headers={"User-Agent": AGENTE})
            with urllib.request.urlopen(pedido, timeout=120) as resposta:
                return cast(bytes, resposta.read())
        except (urllib.error.URLError, TimeoutError) as e:
            erro = e
            time.sleep(pausa * (n + 1))
    raise GeracaoError(f"falha ao baixar {url}: {erro}")


def _b64(consulta: Mapping[str, object]) -> str:
    return base64.b64encode(json.dumps(consulta).encode()).decode()


def _json(bruto: bytes, origem: str) -> Any:
    try:
        return json.loads(bruto)
    except json.JSONDecodeError as erro:
        raise GeracaoError(f"{origem}: resposta não é JSON") from erro


def instrumentos(dia: date, obter: Baixar = baixar) -> tuple[date, bytes]:
    """Cadastro de Instrumentos do último pregão até ``dia`` (volta até 7 dias). Dia sem arquivo
    (feriado, fim de semana) pode vir como erro HTTP 400: tenta o dia anterior."""
    falha: GeracaoError | None = None
    for atras in range(8):
        quando = dia - timedelta(days=atras)
        url = f"{B3_ARQUIVOS}/requestname?fileName=InstrumentsConsolidatedFile&date={quando}"
        try:
            pedido = _json(obter(url), "cadastro de instrumentos")
            token = pedido.get("token") if isinstance(pedido, dict) else None
            if not token:
                continue
            bruto = obter(f"{B3_ARQUIVOS}/?token={token}")
        except GeracaoError as erro:
            log.warning("cadastro de instrumentos de %s: %s", quando, erro)
            falha = erro
            continue
        if bruto.startswith(b"Status do Arquivo"):
            return quando, bruto
    raise GeracaoError(f"cadastro de instrumentos indisponível nos 7 dias antes de {dia}"
                       + (f"; último erro: {falha}" if falha else ""))  # fmt: skip


def empresas(obter: Baixar = baixar, por_pagina: int = 120) -> list[Mapping[str, object]]:
    """Todas as companhias (nome de pregão por código de emissor)."""
    saida: list[Mapping[str, object]] = []
    pagina, total = 1, 1
    while pagina <= total:
        resposta = _json(
            obter(
                f"{B3_EMPRESAS}/GetInitialCompanies/"
                + _b64({"language": "pt-br", "pageNumber": pagina, "pageSize": por_pagina})
            ),
            "empresas",
        )
        try:
            total = int(resposta["page"]["totalPages"])
            saida += list(resposta["results"])
        except (KeyError, TypeError, ValueError) as erro:
            raise GeracaoError(f"empresas: resposta fora do formato ({erro!r})") from erro
        pagina += 1
    return saida


MAX_FALHAS_SEGUIDAS = 20  # a B3 fora do ar: aborta a coleta em vez de esgotar o tempo do job


def eventos(
    emissores: Iterable[str], obter: Baixar = baixar, pausa: float = 0.2, tentativas: int = 2
) -> tuple[dict[str, list[Mapping[str, object]]], list[str]]:
    """Eventos em ações e subscrições de cada emissor (uma consulta por companhia). A B3 às vezes
    devolve resposta vazia: tenta de novo; quem continua sem resposta vai para a lista de falhas
    (os eventos dela já coletados continuam no arquivo, que nunca apaga)."""
    saida: dict[str, list[Mapping[str, object]]] = {}
    falhas: list[str] = []
    seguidas = 0
    for emissor in sorted(set(emissores)):
        url = f"{B3_EMPRESAS}/GetListedSupplementCompany/" + _b64(
            {"issuingCompany": emissor, "language": "pt-br"}
        )
        for n in range(tentativas):
            try:
                resposta = _json(obter(url), f"eventos de {emissor}")
                break
            except GeracaoError:
                time.sleep(pausa * 10 * (n + 1))
        else:
            falhas.append(emissor)
            seguidas += 1
            if seguidas >= MAX_FALHAS_SEGUIDAS:
                raise GeracaoError(f"B3 sem resposta para {MAX_FALHAS_SEGUIDAS} companhias "
                                   "seguidas: coleta abortada")  # fmt: skip
            continue
        seguidas = 0
        if isinstance(resposta, list) and resposta:
            saida[emissor] = cast(list[Mapping[str, object]], resposta)
        time.sleep(pausa)
    return saida, falhas


def _consultar(url: str, origem: str, obter: Baixar, pausa: float, tentativas: int) -> Any:
    """JSON da B3 com novas tentativas (resposta vazia acontece); None se não veio."""
    for n in range(tentativas):
        try:
            return _json(obter(url), origem)
        except GeracaoError:
            time.sleep(pausa * 10 * (n + 1))
    return None


def fundos(
    obter: Baixar = baixar, pausa: float = 0.2, tentativas: int = 2, por_pagina: int = 120
) -> tuple[list[Mapping[str, object]], list[str]]:
    """Fundos listados (FII, Fiagro, ETF, FI-Infra e FIP-IE) com o CNPJ: a lista de cada tipo não
    traz o CNPJ, que vem do detalhe de cada fundo (uma consulta por fundo, como a página de fundos
    da B3). O FIP vem com o tipo da B3 ("FIP IE", "FIP MT"): só o de alíquota zero fica."""
    saida: list[Mapping[str, object]] = []
    falhas: list[str] = []
    seguidas = 0
    for tipo in TIPOS_FUNDO:
        pagina, total, lista = 1, 1, []
        while pagina <= total:
            consulta = {"language": "pt-br", "typeFund": tipo, "pageNumber": pagina,
                        "pageSize": por_pagina}  # fmt: skip
            resposta = _consultar(f"{B3_FUNDOS}/GetListFunds/{_b64(consulta)}", f"fundos {tipo}",
                                  obter, pausa, tentativas)  # fmt: skip
            if resposta is None:
                raise GeracaoError(f"fundos {tipo}: a B3 não respondeu a lista (página {pagina})")
            try:
                total = int(resposta["page"]["totalPages"])
                lista += list(resposta["results"])
            except (KeyError, TypeError, ValueError) as erro:
                raise GeracaoError(f"fundos {tipo}: resposta fora do formato ({erro!r})") from erro
            pagina += 1
        for f in lista:
            sigla = str(f.get("acronym") or "")
            rotulo = tipo
            if tipo == "FIP":
                rotulo = FIP_ALIQUOTA_ZERO.get(str(f.get("typeName") or "").strip(), "")
                if not rotulo:
                    continue  # FIP comum (multiestratégia...): 15%, fica fora até ter regra
            consulta = {"language": "pt-br", "idFNET": str(f.get("id")), "idCEM": sigla,
                        "typeFund": tipo}  # fmt: skip
            detalhe = _consultar(f"{B3_FUNDOS}/GetDetailFund/{_b64(consulta)}", f"fundo {sigla}",
                                 obter, pausa, tentativas)  # fmt: skip
            if not isinstance(detalhe, dict) or not detalhe:
                falhas.append(f"fundo {sigla}")
                seguidas += 1
                if seguidas >= MAX_FALHAS_SEGUIDAS:
                    raise GeracaoError(f"B3 sem resposta para {MAX_FALHAS_SEGUIDAS} fundos "
                                       "seguidos: coleta abortada")  # fmt: skip
                continue
            seguidas = 0
            d = cast(Mapping[str, object], detalhe)
            saida.append({"acronym": sigla, "cnpj": d.get("cnpj"), "typeName": rotulo,
                          "fundName": f.get("fundName"), "tradingName": d.get("tradingName"),
                          "tradingCode": d.get("tradingCode")})  # fmt: skip
            time.sleep(pausa)
    return saida, falhas


def historico(
    nomes: Mapping[str, str], obter: Baixar = baixar, pausa: float = 0.2, tentativas: int = 2,
    por_pagina: int = 120,
) -> tuple[dict[str, list[Mapping[str, object]]], list[str]]:  # fmt: skip
    """Aba "Proventos em dinheiro" de cada companhia (emissor -> nome de pregão), todas as
    páginas: o histórico que GetListedSupplementCompany não traz (lá, só os últimos meses). A
    consulta é pelo nome de pregão; 120 é o maior tamanho de página que a B3 aceita."""
    saida: dict[str, list[Mapping[str, object]]] = {}
    falhas: list[str] = []
    seguidas = 0
    for emissor, nome in sorted(nomes.items()):
        pagina, total, linhas = 1, 1, []
        while pagina <= total:
            consulta = {"language": "pt-br", "pageNumber": pagina, "pageSize": por_pagina,
                        "tradingName": nome}  # fmt: skip
            resposta = _consultar(f"{B3_EMPRESAS}/GetListedCashDividends/{_b64(consulta)}",
                                  f"proventos de {emissor}", obter, pausa, tentativas)  # fmt: skip
            try:
                total = int(resposta["page"]["totalPages"] or 0)
                linhas += list(resposta["results"] or [])
            except (KeyError, TypeError, ValueError):
                resposta = None
            if resposta is None:
                break
            pagina += 1
            time.sleep(pausa)
        else:
            seguidas = 0
            if linhas:
                saida[emissor] = linhas
            continue
        falhas.append(f"proventos de {emissor}")
        seguidas += 1
        if seguidas >= MAX_FALHAS_SEGUIDAS:
            raise GeracaoError(f"B3 sem resposta para {MAX_FALHAS_SEGUIDAS} companhias "
                               "seguidas: coleta abortada")  # fmt: skip
    return saida, falhas


def selic(obter: Baixar = baixar) -> list[Mapping[str, str]]:
    resposta = _json(obter(BCB_SELIC), "Selic")
    if not isinstance(resposta, list):
        raise GeracaoError("Selic: resposta inesperada")
    return cast(list[Mapping[str, str]], resposta)


def sufixos_especificacao(pasta: Path) -> set[str]:
    return {r["sufixo"] for r in ler_csv(pasta / "sufixos_especificacao.csv")}


def nomes_recentes(dia: date, sufixos: Collection[str],
                   obter: Baixar = baixar) -> tuple[list[dict[str, str]], list[str]]:  # fmt: skip
    """Nomes de pregão do COTAHIST do ano de ``dia`` e, em janeiro, do anterior (fecha dezembro; o
    arquivo do ano novo pode nem existir ainda). Falta vai para o resumo: a tabela nunca apaga."""
    saida: list[dict[str, str]] = []
    falhas: list[str] = []
    for ano in [dia.year - 1, dia.year] if dia.month == 1 else [dia.year]:
        try:
            saida += hist.ler_nomes(obter(B3_COTAHIST.format(ano=ano)), sufixos)
        except (GeracaoError, zipfile.BadZipFile) as erro:
            log.warning("COTAHIST de %d: %s", ano, erro)
            falhas.append(f"COTAHIST de {ano}")
    return saida, falhas


def nomes_historicos(ate: date, sufixos: Collection[str], cache: Path | None = None,
                     obter: Baixar = baixar) -> list[dict[str, str]]:  # fmt: skip
    """Nomes de pregão de todos os anos com notas, de 2008 até ``ate`` (uns 1,5 GB de zips)."""
    saida: list[dict[str, str]] = []
    for ano in range(NOMES_DESDE, ate.year + 1):
        saida += hist.ler_nomes(_em_cache(B3_COTAHIST.format(ano=ano), cache, obter), sufixos)
    return saida


def coletar(dia: date, obter: Baixar = baixar, pausa: float = 0.2,
            sufixos: Collection[str] = ()) -> Coleta:  # fmt: skip
    """Uma coleta completa. Eventos e histórico de proventos só das companhias com ação ou unit no
    cadastro."""
    quando, bruto = instrumentos(dia, obter)
    lista, _ = ler_instrumentos(bruto)
    emissores = {i.emissor for i in lista if i.categoria in ("SHARES", "UNIT")}
    log.info("cadastro de %s: %d instrumentos, %d emissores", quando, len(lista), len(emissores))
    companhias = empresas(obter)
    por_emissor, falhas = eventos(emissores, obter, pausa)
    nomes = {e: n for e, n in nomes_de_pregao(companhias).items() if e in emissores}
    por_historico, falhas_historico = historico(nomes, obter, pausa)
    lista_fundos, falhas_fundos = fundos(obter, pausa)
    nomes, falhas_nomes = nomes_recentes(dia, sufixos, obter)
    return Coleta(data=dia, instrumentos=bruto, empresas=companhias, eventos=por_emissor,
                  selic=selic(obter),
                  sem_resposta=tuple(falhas + falhas_historico + falhas_fundos + falhas_nomes),
                  fundos=lista_fundos, historico=por_historico, nomes=nomes)  # fmt: skip


def _em_cache(url: str, cache: Path | None, obter: Baixar) -> bytes:
    """Arquivo anual grande (FRE, COTAHIST): baixa uma vez para ``cache`` e reaproveita."""
    if cache is None:
        return obter(url)
    destino = cache / url.rsplit("/", 1)[1]
    if not destino.is_file():
        cache.mkdir(parents=True, exist_ok=True)
        destino.write_bytes(obter(url))
    return destino.read_bytes()


def historico_eventos(emissor_por_cnpj: Mapping[str, str], ate: date, cache: Path | None = None,
                      obter: Baixar = baixar, pares: Iterable[tuple[str, str]] = (),
                      ) -> tuple[list[EventoB3], list[Revisao], hist.Series,
                                 list[str]]:  # fmt: skip
    """Desdobramentos, grupamentos e bonificações desde 2008: FRE da CVM (2010 a 2023, cada um com
    os três exercícios anteriores) com a data com conferida no COTAHIST da B3 (2007 até ``ate``).
    ``pares`` (ISIN antigo, ISIN novo) das trocas de código encadeiam as séries. A companhia que
    saiu da bolsa entra pelo código que o FCA dá ao CNPJ (``historico.deslistadas``). Devolve
    também as séries de todas as ações e units e o calendário, para a varredura
    (``historico.saltos``)."""
    linhas = [x for ano in ANOS_FRE
              for x in hist.ler_fre(_em_cache(CVM_FRE.format(ano=ano), cache, obter))]  # fmt: skip
    codigos = [hist.ler_fca_codigos(_em_cache(CVM_FCA.format(ano=ano), cache, obter))
               for ano in range(FCA_DESDE, ate.year + 1)]  # fmt: skip
    pares = list(pares)
    series: hist.Series = {}
    pregoes: set[str] = set()
    for ano in range(2007, ate.year + 1):
        hist.ler_cotahist(_em_cache(B3_COTAHIST.format(ano=ano), cache, obter), series, pregoes)
        log.info("COTAHIST %d: %d ISIN", ano, len(series))
    for serie in series.values():
        serie.sort(key=lambda p: p.dia)
    series = hist.encadear(series, pares)
    fora = hist.deslistadas(codigos, emissor_por_cnpj, series)
    log.info("companhias fora do cadastro de hoje, pelo FCA: %d", len(fora))
    eventos, revisao = hist.eventos_fre(linhas, {**emissor_por_cnpj, **fora}, series, pregoes)
    return eventos, revisao, series, sorted(pregoes)
