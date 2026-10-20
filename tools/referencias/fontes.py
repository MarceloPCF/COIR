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
from collections.abc import Callable, Iterable, Mapping
from datetime import date, timedelta
from typing import Any, cast

from tools.referencias.gerar import Coleta, GeracaoError, ler_instrumentos

log = logging.getLogger(__name__)
AGENTE = "COIR/3.0 (tabelas de referencia publicas; github.com/MarceloPCF/COIR)"
B3_ARQUIVOS = "https://arquivos.b3.com.br/api/download"
B3_EMPRESAS = "https://sistemaswebb3-listados.b3.com.br/listedCompaniesProxy/CompanyCall"
B3_FUNDOS = "https://sistemaswebb3-listados.b3.com.br/fundsListedProxy/Search"
TIPOS_FUNDO = ("FII", "FIAGRO", "ETF", "ETF-RF", "ETF-FII")  # os que viram ativo no COIR
BCB_SELIC = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.4390/dados?formato=json"
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
    """Cadastro de Instrumentos do último pregão até ``dia`` (volta até 7 dias)."""
    for atras in range(8):
        quando = dia - timedelta(days=atras)
        url = f"{B3_ARQUIVOS}/requestname?fileName=InstrumentsConsolidatedFile&date={quando}"
        pedido = _json(obter(url), "cadastro de instrumentos")
        token = pedido.get("token") if isinstance(pedido, dict) else None
        if not token:
            continue
        bruto = obter(f"{B3_ARQUIVOS}/?token={token}")
        if bruto.startswith(b"Status do Arquivo"):
            return quando, bruto
    raise GeracaoError(f"cadastro de instrumentos indisponível nos 7 dias antes de {dia}")


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
    """Fundos listados (FII, Fiagro e ETF) com o CNPJ: a lista de cada tipo não traz o CNPJ, que
    vem do detalhe de cada fundo (uma consulta por fundo, como a página de fundos da B3)."""
    saida: list[Mapping[str, object]] = []
    falhas: list[str] = []
    seguidas = 0
    for tipo in TIPOS_FUNDO:
        pagina, total, lista = 1, 1, []
        while pagina <= total:
            consulta = {"language": "pt-br", "typeFund": tipo, "pageNumber": pagina,
                        "pageSize": por_pagina}  # fmt: skip
            resposta = _json(obter(f"{B3_FUNDOS}/GetListFunds/{_b64(consulta)}"), f"fundos {tipo}")
            try:
                total = int(resposta["page"]["totalPages"])
                lista += list(resposta["results"])
            except (KeyError, TypeError, ValueError) as erro:
                raise GeracaoError(f"fundos {tipo}: resposta fora do formato ({erro!r})") from erro
            pagina += 1
        for f in lista:
            sigla = str(f.get("acronym") or "")
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
            saida.append({"acronym": sigla, "cnpj": d.get("cnpj"), "typeName": tipo,
                          "fundName": f.get("fundName")})  # fmt: skip
            time.sleep(pausa)
    return saida, falhas


def selic(obter: Baixar = baixar) -> list[Mapping[str, str]]:
    resposta = _json(obter(BCB_SELIC), "Selic")
    if not isinstance(resposta, list):
        raise GeracaoError("Selic: resposta inesperada")
    return cast(list[Mapping[str, str]], resposta)


def coletar(dia: date, obter: Baixar = baixar, pausa: float = 0.2) -> Coleta:
    """Uma coleta completa. Eventos só das companhias com ação ou unit no cadastro."""
    quando, bruto = instrumentos(dia, obter)
    lista, _ = ler_instrumentos(bruto)
    emissores = {i.emissor for i in lista if i.categoria in ("SHARES", "UNIT")}
    log.info("cadastro de %s: %d instrumentos, %d emissores", quando, len(lista), len(emissores))
    por_emissor, falhas = eventos(emissores, obter, pausa)
    lista_fundos, falhas_fundos = fundos(obter, pausa)
    return Coleta(data=dia, instrumentos=bruto, empresas=empresas(obter), eventos=por_emissor,
                  selic=selic(obter), sem_resposta=tuple(falhas + falhas_fundos),
                  fundos=lista_fundos)  # fmt: skip
