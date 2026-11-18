"""Cadastros da CVM (dados abertos, ODbL): códigos de negociação das companhias e cotas de FII e
Fiagro negociadas em bolsa, com o CNPJ e o nome.

- ``valores_mobiliarios_cvm.csv``: do FCA (formulário cadastral, tabela ``valor_mobiliario``, de
  2018 em diante), cada código de negociação de ação ou unit em bolsa, com o CNPJ da companhia, a
  razão social (a do cadastro de companhias abertas, ou a do FCA), a espécie e o último ano em que
  o FCA o trouxe;
- ``fundos_cvm.csv``: dos informes mensais de FII (2016 em diante) e de Fiagro (os doze meses que
  a CVM mantém), o ISIN da cota negociada em bolsa, com o CNPJ, o nome e o tipo.

As tabelas nunca apagam: o FCA e o informe só trazem o que a companhia ou o fundo declara agora, e
o Bens e Direitos de anos passados usa o que valia na época. A linha mais recente vale.
"""

from __future__ import annotations

import csv
import io
import re
import zipfile
from collections import defaultdict
from collections.abc import Collection, Iterable, Mapping
from typing import Any

COLUNAS_VALORES = ("ticker", "cnpj", "razao_social", "especie", "visto_ate")
COLUNAS_FUNDOS = ("isin", "cnpj", "nome", "tipo", "visto_ate")
VALORES = "valores_mobiliarios_cvm.csv"
FUNDOS = "fundos_cvm.csv"
ESPECIES = {"Ações Ordinárias": "ON", "Ações Preferenciais": "PN", "Units": "UNIT"}
RE_CODIGO = re.compile(r"[A-Z][A-Z0-9]{3}([3-8]|11)")
RE_ISIN_COTA = re.compile(
    r"BR[A-Z0-9]{4}CTF[0-9A-Z]{2}\d"
)  # cota de fundo (CTF); recibo e direito não


def _so_digitos(texto: str) -> str:
    return re.sub(r"\D", "", texto)


def _ler(texto: str) -> list[dict[str, str]]:
    return list(csv.DictReader(io.StringIO(texto), delimiter=";"))


def _do_zip(bruto: bytes, trecho: str, sem: str = "") -> list[dict[str, str]]:
    """O CSV do zip anual cujo nome tem ``trecho`` (e não tem ``sem``); vazio se não houver."""
    with zipfile.ZipFile(io.BytesIO(bruto)) as z:
        nome = next((n for n in z.namelist() if trecho in n and not (sem and sem in n)), None)
        return _ler(z.read(nome).decode("latin-1")) if nome else []


def razoes_sociais(cadastro: bytes) -> dict[str, str]:
    """CNPJ -> razão social, do cadastro de companhias abertas (``cad_cia_aberta.csv``)."""
    return {_so_digitos(r["CNPJ_CIA"]): r["DENOM_SOCIAL"].strip()
            for r in _ler(cadastro.decode("latin-1")) if r.get("DENOM_SOCIAL")}  # fmt: skip


def ler_valores(zips_fca: Iterable[bytes], razoes: Mapping[str, str]) -> list[dict[str, str]]:
    """Códigos de ação e unit negociados em bolsa, um por ticker (a declaração mais recente)."""
    saida: dict[str, dict[str, str]] = {}
    for bruto in zips_fca:
        for r in _do_zip(bruto, "valor_mobiliario"):
            codigo = (r.get("Codigo_Negociacao") or "").strip().upper()
            especie = ESPECIES.get((r.get("Valor_Mobiliario") or "").strip())
            if (r.get("Mercado") or "").strip() != "Bolsa" or especie is None:
                continue
            if not RE_CODIGO.fullmatch(codigo):
                continue
            cnpj = _so_digitos(r["CNPJ_Companhia"])
            ano = (r.get("Data_Referencia") or "")[:10]
            if ano < saida.get(codigo, {}).get("visto_ate", ""):
                continue
            saida[codigo] = {"ticker": codigo, "cnpj": cnpj,
                             "razao_social": razoes.get(cnpj) or r["Nome_Empresarial"].strip(),
                             "especie": especie, "visto_ate": ano}  # fmt: skip
    return [saida[t] for t in sorted(saida)]


def ler_fundos(zips_fii: Iterable[bytes], zips_fiagro: Iterable[bytes],
               ) -> tuple[list[dict[str, str]], set[str]]:  # fmt: skip
    """Cotas de FII e Fiagro negociadas em bolsa, uma por ISIN (o informe mais recente), e os ISIN
    ambíguos. O ISIN é declarado pelo próprio fundo: o que dois CNPJ declararam (um fundo novo que
    repetiu o ISIN de outro) fica de fora, porque não se sabe qual dos dois é o certo."""
    saida: dict[str, dict[str, str]] = {}
    declarantes: dict[str, set[str]] = defaultdict(set)

    def incluir(isin: str, cnpj: str, nome: str, tipo: str, quando: str) -> None:
        isin = isin.strip().upper()
        if not RE_ISIN_COTA.fullmatch(isin):
            return
        declarantes[isin].add(_so_digitos(cnpj))
        if quando < saida.get(isin, {}).get("visto_ate", ""):
            return
        saida[isin] = {"isin": isin, "cnpj": _so_digitos(cnpj), "nome": nome.strip(),
                       "tipo": tipo, "visto_ate": quando}  # fmt: skip

    for bruto in zips_fii:
        for r in _do_zip(bruto, "geral"):
            if (r.get("Mercado_Negociacao_Bolsa") or "").strip() == "S":
                # antes da RCVM 175 as colunas eram CNPJ_Fundo e Nome_Fundo
                incluir(r["Codigo_ISIN"], r.get("CNPJ_Fundo_Classe") or r["CNPJ_Fundo"],
                        r.get("Nome_Fundo_Classe") or r["Nome_Fundo"], "FII",
                        r["Data_Referencia"][:10])  # fmt: skip
    for bruto in zips_fiagro:
        for r in _do_zip(bruto, "fiagro", sem="subclasse"):
            if (r.get("Mercado_Negociacao") or "").strip().upper() == "BOLSA":
                incluir(r["Codigo_ISIN"], r["CNPJ_Classe"], r["Nome_Classe"], "FIAGRO",
                        r["Data_Referencia"][:10])  # fmt: skip
    ambiguos = {i for i, cnpjs in declarantes.items() if len(cnpjs) > 1}
    return [saida[i] for i in sorted(saida) if i not in ambiguos], ambiguos


def juntar(anteriores: Iterable[Mapping[str, str]], novas: Iterable[Mapping[str, str]],
           chave: str, fora: Collection[str] = ()) -> list[dict[str, str]]:  # fmt: skip
    """A tabela só apaga o que ficou ambíguo (``fora``): o que a coleta nova traz substitui a
    linha da mesma chave quando é mais recente; o que ela não traz continua (o Fiagro só tem os
    últimos doze meses)."""
    todas: dict[str, dict[str, str]] = {
        r[chave]: dict(r) for r in anteriores if r[chave] not in fora
    }
    for r in novas:
        if r["visto_ate"] >= todas.get(r[chave], {}).get("visto_ate", ""):
            todas[r[chave]] = dict(r)
    return [todas[k] for k in sorted(todas) if k not in fora]


def contagens(valores: list[dict[str, str]], fundos: list[dict[str, str]],
              ambiguos: Collection[str]) -> dict[str, Any]:  # fmt: skip
    """Resumo da coleta (para o texto do comando)."""
    return {"tickers": len(valores), "companhias": len({r["cnpj"] for r in valores}),
            "fundos": len({r["cnpj"] for r in fundos}), "cotas": len(fundos),
            "ambiguos": len(ambiguos)}  # fmt: skip
