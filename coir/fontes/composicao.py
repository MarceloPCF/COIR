"""Junta os provedores por conceito, com a precedência curadoria > CVM/BCB > bolsa.

O mesmo registro em duas fontes entra uma vez, com as origens somadas. Conteúdo diferente na mesma
chave é divergência: vale a fonte de maior precedência e a diferença fica registrada (a mesma regra
de ``juntar_eventos`` do pipeline), nunca uma escolha calada.
"""

import logging
import re
from collections.abc import Iterable
from dataclasses import replace
from datetime import date
from decimal import Decimal
from functools import cache
from pathlib import Path

from coir.fontes.modelo import Divergencia, Evento, partes_da_origem
from coir.fontes.provedores import b3, curadoria, cvm

log = logging.getLogger(__name__)

RE_TICKER_PADRAO = re.compile(r"^[A-Z0-9]{4}([3-8]|11|5[3-6])$")  # 53 a 56: ações "P" (RAIZ55)


def _somar(atual: Evento, outro: Evento) -> Evento:
    """O mesmo evento visto por outra fonte: soma a origem e as referências por documento."""
    origens = partes_da_origem(atual.origem) | partes_da_origem(outro.origem)
    refs = [r for texto in (atual.referencia, outro.referencia) for r in texto.split("; ") if r]
    return replace(atual, origem="+".join(sorted(origens)),
                   referencia="; ".join(dict.fromkeys(refs)),
                   custo_atribuido=atual.custo_atribuido or outro.custo_atribuido)  # fmt: skip


def _mesmo_conteudo(a: Evento, b: Evento) -> bool:
    return (replace(a, origem="", referencia="", custo_atribuido="")
            == replace(b, origem="", referencia="", custo_atribuido="")
            and not (a.custo_atribuido and b.custo_atribuido
                     and a.custo_atribuido != b.custo_atribuido))  # fmt: skip


def juntar_eventos(*por_precedencia: Iterable[Evento]) -> tuple[list[Evento], list[Divergencia]]:
    """Eventos das fontes (a primeira lista vale mais), na ordem da chave (a mesma em que o
    pipeline grava a tabela) e sem os pares que se anulam (``_neutros``). A chave é a do pipeline
    (``COLUNAS_CHAVE`` em ``tools/referencias/gerar.py``), única por construção dentro de cada
    fonte: o que repete a chave com outro conteúdo só pode vir de outra fonte (divergência)."""
    escolhidos: dict[tuple[str, ...], Evento] = {}
    divergencias: list[Divergencia] = []
    for lista in por_precedencia:
        for ev in lista:
            atual = escolhidos.get(ev.chave)
            if atual is None:
                escolhidos[ev.chave] = ev
            elif _mesmo_conteudo(atual, ev):
                escolhidos[ev.chave] = _somar(atual, ev)
            else:
                divergencias.append(Divergencia("evento", ev.chave, atual.origem, ev.origem))
    ordenados = [escolhidos[k] for k in sorted(escolhidos)]
    neutros = _neutros(ordenados)
    return [e for i, e in enumerate(ordenados) if i not in neutros], divergencias


def _neutros(eventos: list[Evento]) -> set[int]:
    """Grupamento e desdobramento no mesmo ISIN e data com que se anulam (ITUB 2011: 100 para 1 e
    1 para 100): a quantidade não muda, então não entram em controle nenhum. Ficam no pacote (que
    nunca apaga): se a fonte voltar a listar só um dos lados, a coleta não o toma por evento
    novo."""
    grupos: dict[tuple[str, str, str], list[int]] = {}
    for i, e in enumerate(eventos):
        if e.tipo in ("GRUPAMENTO", "DESDOBRAMENTO", "BONIFICACAO"):
            grupos.setdefault((e.isin, e.data_com, e.isin_emitido), []).append(i)
    neutros: set[int] = set()
    for indices in grupos.values():
        grupamento = {eventos[i].tipo == "GRUPAMENTO" for i in indices}
        produto = Decimal(1)
        for i in indices:
            produto *= eventos[i].multiplicador
        if grupamento == {True, False} and produto == 1:
            neutros.update(indices)
    return neutros


def eventos_e_divergencias(pasta: Path) -> tuple[list[Evento], list[Divergencia]]:
    """Os eventos e as divergências entre as fontes, que ficam no log (chamada explícita, para
    conferência: as propostas usam ``eventos``, várias vezes, sem repetir o aviso)."""
    lista, divergencias = juntar_eventos(cvm.eventos(pasta), b3.eventos(pasta))
    for d in divergencias:
        log.warning("evento %s: %s mantida, %s divergente", d.chave, d.mantida, d.descartada)
    return lista, divergencias


def eventos(pasta: Path) -> list[Evento]:
    """Eventos corporativos utilizáveis de todas as fontes (vazio sem nenhuma)."""
    return juntar_eventos(cvm.eventos(pasta), b3.eventos(pasta))[0]


# classe no ISIN (posições 7 a 11) -> final do código de negociação; confere com todo o cadastro
# de instrumentos de hoje (ON 3, PN 4, PNA 5... unit e cota de fundo 11)
FINAL_PELA_CLASSE = {"ACNOR": "3", "ACNPR": "4", "ACNPA": "5", "ACNPB": "6", "ACNPC": "7",
                     "ACNPD": "8"}  # fmt: skip


def ticker_pelo_isin(isin: str) -> str | None:
    """O código de negociação que a estrutura do ISIN dá (emissor + final da classe), ou None."""
    if len(isin) != 12 or not isin.startswith("BR"):
        return None
    final = FINAL_PELA_CLASSE.get(isin[6:11])
    if final is None and (isin[6:10] == "CDAM" or isin[6:9] == "CTF"):
        final = "11"
    return f"{isin[2:6]}{final}" if final else None


def tickers_por_isin(pasta: Path) -> dict[str, str]:
    """ISIN -> ticker de ação ou unit (o mais recente visto no cadastro de instrumentos). O ISIN de
    antes de uma troca de código (KROT3 era BRKROTACNOR9) leva ao ticker da época, que o cadastro
    de hoje não tem: o evento daquele ISIN acha a posição do ativo. O ISIN de evento que nenhuma
    tabela liga a um ticker vai pela estrutura do código, se o ticker que ela dá existe no
    cadastro da CVM ou nas tabelas curadas. O código reaproveitado não engana: a posição é a do
    ticker na data com do evento, quando o código ainda era de quem fez o evento."""
    padrao = RE_TICKER_PADRAO
    saida: dict[str, tuple[str, str]] = {}
    for i in b3.instrumentos(pasta):
        if padrao.match(i.ticker) and i.visto_ate >= saida.get(i.isin, ("", ""))[0]:
            saida[i.isin] = (i.visto_ate, i.ticker)
    por_isin = {isin: t for isin, (_, t) in saida.items()}
    for t in curadoria.trocas(pasta):
        for isin, ticker in ((t.isin_antigo, t.ticker_impresso), (t.isin_novo, t.ticker)):
            if isin and padrao.match(ticker):
                por_isin.setdefault(isin, ticker)
    conhecidos = cvm.tickers(pasta) | curadoria.tickers(pasta)
    for ev in eventos(pasta):
        for isin in (ev.isin, ev.isin_emitido):
            ticker = ticker_pelo_isin(isin) if isin and isin not in por_isin else None
            if ticker in conhecidos:
                por_isin[isin] = ticker
    return por_isin


def emissores(pasta: Path) -> dict[str, tuple[str, str, str]]:
    """Emissor (4 letras) -> (CNPJ, razão social, tipo: COMPANHIA, FII, FIAGRO, ETF...), para o
    Bens e Direitos. A curadoria vale sobre tudo (o tipo continua o do cadastro). Entre a bolsa e
    a CVM vale a bolsa, que diz quem tem o código hoje (o CNPJ visto por último); a CVM completa
    o código que a bolsa não traz (o código é reaproveitado, e a CVM guarda quem o usou antes)."""
    saida: dict[str, tuple[str, str, str, str]] = {}
    for e in b3.emissores(pasta):
        if e.visto_ate >= saida.get(e.emissor, ("", "", "", ""))[3]:
            saida[e.emissor] = (e.cnpj, e.razao_social, e.tipo, e.visto_ate)
    for e in cvm.emissores(pasta):
        saida.setdefault(e.emissor, (e.cnpj, e.razao_social, e.tipo, e.visto_ate))
    for e in curadoria.emissores(pasta):
        tipo = saida.get(e.emissor, ("", "", "", ""))[2]
        saida[e.emissor] = (e.cnpj, e.razao_social, tipo, "")
    return {e: (c, n, t) for e, (c, n, t, _) in saida.items()}


def instrumentos(pasta: Path) -> dict[str, tuple[str, str, str]]:
    """Ticker -> (emissor, especificação "ON NM", razão social), o visto por último."""
    saida: dict[str, tuple[str, str, str, str]] = {}
    for i in b3.instrumentos(pasta):
        if i.visto_ate >= saida.get(i.ticker, ("", "", "", ""))[3]:
            saida[i.ticker] = (i.emissor, i.especificacao, i.nome, i.visto_ate)
    return {t: (e, s, n) for t, (e, s, n, _) in saida.items()}


@cache
def ultimos_negocios(pasta: Path) -> dict[str, date]:
    """Ticker -> último negócio visto (vazio em pacote antigo)."""
    saida: dict[str, date] = {}
    for n in b3.nomes_pregao(pasta):
        ate = date.fromisoformat(n.ate)
        saida[n.ticker] = max(ate, saida.get(n.ticker, ate))
    return saida


@cache
def saidas(pasta: Path) -> dict[str, dict[str, str]]:
    """O que houve com o ativo que parou de negociar (curada, com fonte)."""
    return curadoria.saidas(pasta)
