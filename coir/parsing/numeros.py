"""Números, datas e rótulos como aparecem no texto das notas SINACOR."""

import re
import unicodedata
from datetime import date, datetime
from decimal import Decimal

from coir.domain.nota import DC, Montante
from coir.erros import ExtracaoError

# dinheiro: 1.234,56 | 1234,56 | -2 casas ou mais (preço BM&F 97.810,0000)
RE_DECIMAL = r"-?(?:\d{1,3}(?:\.\d{3})+|\d+),\d+"
RE_INTEIRO = r"-?(?:\d{1,3}(?:\.\d{3})+|\d+)"
# valor seguido de "|" e/ou D/C opcionais: "44,00 | C", "0,00|D", "0,00 |", "1,49 D"
RE_MONTANTE = re.compile(rf"(?P<valor>{RE_DECIMAL})(?:\s*\|)?(?:\s*(?P<dc>[DC])(?![\w,]))?")
RE_DATA = re.compile(r"\d{2}/\d{2}/\d{4}")


def decimal_br(texto: str) -> Decimal:
    """'1.234,56' -> Decimal('1234.56'); '1800' -> Decimal('1800'); aceita sinal."""
    limpo = texto.strip().replace(".", "").replace(",", ".")
    if not re.fullmatch(r"-?\d+(?:\.\d+)?", limpo):
        raise ExtracaoError(f"número inválido: {texto!r}")
    return Decimal(limpo)


def montante(texto: str) -> Montante:
    m = RE_MONTANTE.fullmatch(texto.strip())
    if not m:
        raise ExtracaoError(f"valor inválido: {texto!r}")
    return Montante(valor=decimal_br(m["valor"]), dc=_dc(m["dc"]))


def montantes(texto: str) -> list[Montante]:
    """Todos os valores (com D/C) de um trecho, na ordem."""
    return [
        Montante(valor=decimal_br(m["valor"]), dc=_dc(m["dc"])) for m in RE_MONTANTE.finditer(texto)
    ]


def data_br(texto: str) -> date:
    try:
        return datetime.strptime(texto.strip(), "%d/%m/%Y").date()
    except ValueError as erro:
        raise ExtracaoError(f"data inválida: {texto!r}") from erro


def _dc(texto: str | None) -> DC | None:
    if texto == "D":
        return "D"
    if texto == "C":
        return "C"
    return None


def compactar(texto: str) -> str:
    """Forma de comparação de rótulos: sem acentos, sem espaços, minúsculas.

    O layout antigo da XP cola palavras ("Taxa deliquidação", "TaxaOperacional"); comparar a forma
    compacta torna os rótulos equivalentes.
    """
    sem_acento = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", "", sem_acento).casefold()


def regex_rotulo(rotulo: str) -> re.Pattern[str]:
    """Regex que casa o rótulo com espaços opcionais entre os caracteres e acentos trocados."""
    partes: list[str] = []
    for ch in compactar(rotulo):
        base = re.escape(ch)
        partes.append(f"[{base}{_ACENTOS.get(ch, '')}]" if ch in _ACENTOS else base)
    return re.compile(r"\s*".join(partes), re.IGNORECASE)


_ACENTOS = {"a": "áàâãÁÀÂÃ", "e": "éêÉÊ", "i": "íÍ", "o": "óôõÓÔÕ", "u": "úÚü", "c": "çÇ"}


def valor_apos_rotulo(texto: str, rotulo: str) -> Montante | None:
    """Primeiro valor que segue o rótulo na mesma linha (ex.: 'Taxa de liquidação 1,49 D')."""
    padrao = regex_rotulo(rotulo)
    for linha in texto.splitlines():
        m = padrao.search(linha)
        if not m:
            continue
        valor = RE_MONTANTE.match(linha[m.end() :].lstrip())
        if valor:
            return Montante(valor=decimal_br(valor["valor"]), dc=_dc(valor["dc"]))
    return None
