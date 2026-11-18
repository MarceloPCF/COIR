"""Dados abertos da CVM: eventos do formulário de referência (FRE) e dos avisos aos acionistas
(IPE), com o protocolo do documento na referência; códigos de negociação das companhias (FCA) e
cotas de FII e Fiagro em bolsa (informes mensais), com o CNPJ.

O código de emissor (as 4 letras) volta a ser usado pela bolsa depois que a companhia ou o fundo
sai: o código de um FII de 2019 hoje pode ser de um ETF. Por isso o cadastro da CVM só completa o
que a bolsa não traz, e o código que a CVM liga a mais de um CNPJ (ou a uma companhia e a um
fundo) fica de fora.
"""

import re
from collections import defaultdict
from pathlib import Path

from coir.domain.cpf import cnpj_valido
from coir.fontes.leitura import opcionais
from coir.fontes.modelo import (
    CVM_FCA,
    CVM_FUNDOS,
    ORIGENS_CVM,
    Emissor,
    Evento,
    evento_valido,
    partes_da_origem,
)
from coir.fontes.provedores import b3

EVENTOS = "eventos_cvm.csv"  # formato 3: só as linhas da CVM, com conferido_cotahist
EVENTOS_JUNTOS = "eventos_b3.csv"  # formato 2: as origens juntas numa linha (coluna origem)
VALORES = "valores_mobiliarios_cvm.csv"
FUNDOS = "fundos_cvm.csv"
RE_EMISSOR = re.compile(r"[A-Z][A-Z0-9]{3}")  # código de emissor de verdade começa por letra


def eventos(pasta: Path) -> list[Evento]:
    """Eventos utilizáveis com alguma origem na CVM: o arquivo da CVM e, com a fonte da bolsa
    ligada, as linhas da CVM que estiverem na tabela junta (pacote do formato 2). Desligada a
    fonte da bolsa, a tabela junta não é aberta."""
    linhas = opcionais(pasta / EVENTOS)
    if b3.ligada():
        linhas += opcionais(pasta / EVENTOS_JUNTOS)
    return [Evento.de_linha(r) for r in linhas
            if partes_da_origem(r.get("origem", "")) & ORIGENS_CVM
            and evento_valido(r)]  # fmt: skip


def _emissores_por_codigo(pasta: Path) -> dict[str, list[Emissor]]:
    """Código -> emissores que a CVM liga a ele (companhia pelo ticker, fundo pelo ISIN)."""
    saida: dict[str, list[Emissor]] = defaultdict(list)
    for r in opcionais(pasta / VALORES):
        codigo = r["ticker"][:4]
        saida[codigo].append(Emissor(emissor=codigo, cnpj=r["cnpj"],
                                     razao_social=r["razao_social"], tipo="COMPANHIA",
                                     visto_ate=r["visto_ate"], origem=CVM_FCA))  # fmt: skip
    for r in opcionais(pasta / FUNDOS):
        codigo = r["isin"][2:6]
        saida[codigo].append(Emissor(emissor=codigo, cnpj=r["cnpj"], razao_social=r["nome"],
                                     tipo=r["tipo"], visto_ate=r["visto_ate"],
                                     origem=CVM_FUNDOS))  # fmt: skip
    return saida


def emissores(pasta: Path) -> list[Emissor]:
    """Um emissor por código, só quando a CVM liga o código a um único CNPJ válido; o fundo que
    trocou de classe (ISIN novo, mesmo CNPJ) fica com o informe mais recente."""
    saida: list[Emissor] = []
    for codigo, lista in sorted(_emissores_por_codigo(pasta).items()):
        if not RE_EMISSOR.fullmatch(codigo) or len({e.cnpj for e in lista}) != 1:
            continue
        if len({e.tipo == "COMPANHIA" for e in lista}) != 1 or not cnpj_valido(lista[0].cnpj):
            continue
        saida.append(max(lista, key=lambda e: (e.visto_ate, e.razao_social)))
    return saida


def tipos_ativo(pasta: Path) -> list[tuple[str, str]]:
    """Ticker -> tipo que o cadastro da CVM garante: a cota de FII ou Fiagro (código + 11; o
    Fiagro é FII para o imposto) e a unit da companhia. Ação ON e PN já têm o tipo pelo
    código.

    A cota negociada é sempre o código do ISIN + 11 (confere com todo o cadastro de instrumentos
    da bolsa). Limite conhecido: com a fonte da bolsa desligada, o código de um FII antigo que a
    bolsa deu depois a um ETF (FIXX11) sai como FII; ligada, vale o tipo da bolsa."""
    saida: list[tuple[str, str]] = []
    lista = emissores(pasta)
    for e in lista:
        if e.tipo in ("FII", "FIAGRO"):
            saida.append((f"{e.emissor}11", "FII"))
    unicos = {e.emissor for e in lista}
    for r in opcionais(pasta / VALORES):
        if r["especie"] == "UNIT" and r["ticker"][:4] in unicos:
            saida.append((r["ticker"], "UNIT"))
    return sorted(saida)


def tickers(pasta: Path) -> frozenset[str]:
    """Tickers de ação, unit e cota de fundo que o cadastro da CVM confirma (sem data: a posição
    é procurada pelo ticker na data do evento, quando o código era de quem fez o evento)."""
    return frozenset({r["ticker"] for r in opcionais(pasta / VALORES)}
                     | {t for t, _ in tipos_ativo(pasta)})  # fmt: skip


def companhias_por_ticker(pasta: Path) -> dict[str, Emissor]:
    """Ticker de ação ou unit -> a companhia que o declarou no FCA (o CNPJ do registro na CVM),
    com o ano da última declaração em ``visto_ate``."""
    return {r["ticker"]: Emissor(emissor=r["ticker"][:4], cnpj=r["cnpj"],
                                 razao_social=r["razao_social"], tipo="COMPANHIA",
                                 visto_ate=r["visto_ate"], origem=CVM_FCA)
            for r in opcionais(pasta / VALORES) if cnpj_valido(r["cnpj"])}  # fmt: skip
