"""Rótulos das notas SINACOR: as formas aceitas de cada um ficam em rotulos.yaml.

Cada rótulo é procurado pela chave ("data_pregao"), nunca por um texto solto no código: a
tolerância (acento, maiúsculas, espaço opcional entre as palavras, abreviações, "do/da/de"
opcional, quebra de linha entre as palavras) vem da compilação das formas do YAML numa regex só,
igual para todas as gramáticas. A sintaxe das formas está no cabeçalho do YAML.

A regex casa só o rótulo, sem ignorar caixa no que vem depois: quem a compõe com o valor
(``padrao``) continua lendo "D"/"C" em maiúscula.
"""

import re
import unicodedata
from collections.abc import Sequence
from functools import cache
from importlib import resources

import yaml

from coir.domain.nota import Montante
from coir.erros import ConfiguracaoError, ExtracaoError
from coir.parsing.numeros import RE_MONTANTE, montante

# o rótulo não começa no meio de uma palavra (vale qualquer letra, com acento ou não)
_INICIO = r"(?<![^\W\d_])"
_NAO_CONTINUA = r"(?![^\W\d_])"
_DENTRO = r"[^\S\n]*"  # espaço solto dentro da palavra, sem passar de linha
_ENTRE = r"\s*"  # entre as palavras: nenhum, vários ou uma quebra de linha


def _sem_acento(c: str) -> str:
    return unicodedata.normalize("NFKD", c).encode("ascii", "ignore").decode().casefold()


_ACENTOS = {
    b: b + "".join(c for c in "áàâãäéèêëíìîïóòôõöúùûüç" if _sem_acento(c) == b) for b in "aeiouc"
}


@cache
def _formas() -> dict[str, tuple[str, ...]]:
    arquivo = resources.files("coir.parsing").joinpath("rotulos.yaml")
    dados: object = yaml.safe_load(arquivo.read_text(encoding="utf-8"))
    if not isinstance(dados, dict):
        raise ConfiguracaoError("rotulos.yaml: esperado um mapa de rótulo para formas")
    formas: dict[str, tuple[str, ...]] = {}
    for chave, lista in dados.items():  # pyright: ignore[reportUnknownVariableType]
        if not (
            isinstance(chave, str)
            and isinstance(lista, list)
            and lista
            and all(isinstance(f, str) and f.strip() for f in lista)  # pyright: ignore[reportUnknownVariableType]
        ):
            raise ConfiguracaoError(f"rotulos.yaml: {chave!r} precisa de uma lista de formas")
        formas[chave] = tuple(str(f) for f in lista)  # pyright: ignore[reportUnknownArgumentType]
    return formas


def chaves() -> tuple[str, ...]:
    return tuple(_formas())


def formas(chave: str) -> tuple[str, ...]:
    """As formas aceitas do rótulo, na ordem do YAML (a primeira é a de referência)."""
    try:
        return _formas()[chave]
    except KeyError as erro:
        raise ConfiguracaoError(f"rótulo {chave!r} ausente de rotulos.yaml") from erro


def _caractere(c: str) -> str:
    if c in "º°":
        return "[º°]"
    base = _sem_acento(c)
    if len(base) == 1 and base in _ACENTOS:
        return f"[{_ACENTOS[base]}]"
    return re.escape(c)


def _palavra(palavra: str) -> str:
    prefixo = palavra.endswith("*")
    palavra = palavra.removesuffix("*")
    abreviacao = not prefixo and palavra.endswith(".") and len(palavra) > 1
    palavra = palavra.removesuffix(".") if abreviacao else palavra
    # o ponto do meio ("A.N.A.") fica preso à letra de antes e só leva os espaços se houver ponto:
    # "\.?" solto entre dois "[ ]*" daria retrocesso explosivo em linha cheia de espaços
    pedacos: list[str] = []
    for c in palavra:
        if c == "." and pedacos:
            pedacos[-1] += rf"(?:{_DENTRO}\.)?"
        else:
            pedacos.append(_caractere(c))
    corpo = _DENTRO.join(pedacos)
    if prefixo:
        return corpo + r"\S*"
    if not abreviacao:
        return corpo
    # letra solta ("D.") só com o ponto: sem ele seria um D/C ou uma inicial qualquer
    return corpo + (r"\." if len(palavra) == 1 else rf"(?:\.|{_NAO_CONTINUA})")


def _token(token: str) -> str:
    opcional = token.startswith("[") and token.endswith("]")
    alternativas = "|".join(_palavra(p) for p in token.strip("[]").split("|") if p)
    return f"(?:{alternativas})?" if opcional else f"(?:{alternativas})"


def _forma(forma: str) -> str:
    return _ENTRE.join(_token(t) for t in forma.split())


def _alternativas(chave: str) -> str:
    regexes: list[str] = []
    for forma in formas(chave):
        texto = _forma(forma)
        if re.fullmatch(texto, "", re.IGNORECASE):  # só palavras opcionais: casaria qualquer linha
            raise ConfiguracaoError(f"rotulos.yaml: {chave!r} tem forma vazia: {forma!r}")
        regexes.append(texto)
    return "|".join(regexes)


@cache
def padrao(*chaves: str, colado: bool = False) -> str:
    """Regex (texto) dos rótulos em sequência, ignorando caixa só neles: para compor com o valor.
    ``colado``: o rótulo pode vir grudado na palavra anterior (a XP imprime "IRRFDay Trade"); vale
    também para as guardas, que precisam ver o rótulo mesmo colado para falhar alto.

    Entre as palavras, o espaço inclui a quebra de linha: aplicada ao texto inteiro da página, a
    regex aceita o rótulo quebrado em qualquer linha (``procurar`` prefere a linha inteira)."""
    if not chaves:
        raise ConfiguracaoError("padrao() sem rótulo")
    inicio = "" if colado else _INICIO
    return "(?i:" + inicio + _ENTRE.join(f"(?:{_alternativas(c)})" for c in chaves) + ")"


@cache
def regex(*chaves: str, colado: bool = False) -> re.Pattern[str]:
    return re.compile(padrao(*chaves, colado=colado))


def referencia(*chaves: str) -> str:
    """Como o rótulo aparece nas mensagens: a primeira forma de cada chave, sem a sintaxe."""

    def limpa(forma: str) -> str:
        palavras = [t.strip("[]").split("|")[0].removesuffix("*") for t in forma.split()]
        # "/CCP" é o fim da palavra anterior ("Taxa de liquidação/CCP")
        return re.sub(r" (/\w)", r"\1", " ".join(palavras))

    return " ".join(limpa(formas(c)[0]) for c in chaves)


def procurar(linhas: Sequence[str], *chaves: str, inicio: int = 0) -> tuple[int, int]:
    """Primeira e última linha do primeiro rótulo a partir de ``inicio``. Só quando nenhuma linha
    o traz inteiro, vale o rótulo quebrado entre duas linhas seguidas ("Data" / "pregão")."""
    alvo = regex(*chaves)
    for i in range(inicio, len(linhas)):
        if alvo.search(linhas[i]):
            return i, i
    for i in range(inicio, len(linhas) - 1):
        if alvo.search(f"{linhas[i]}\n{linhas[i + 1]}"):
            return i, i + 1
    raise ExtracaoError(f"rótulo {referencia(*chaves)!r} não encontrado")


def no_inicio(texto: str, *chaves: str) -> bool:
    """O texto começa pelos rótulos (espaços do começo ignorados)."""
    return regex(*chaves).match(texto.lstrip()) is not None


def linha_inteira(texto: str, *chaves: str) -> bool:
    """O texto é só o rótulo (espaços das pontas ignorados)."""
    return regex(*chaves).fullmatch(texto.strip()) is not None


def valor_apos(texto: str, chave: str) -> Montante | None:
    """Primeiro valor logo depois do rótulo, na mesma linha ('Taxa de liquidação 1,49 D'). Linha
    em que o rótulo vem sem valor (título de seção) não conta: vale a próxima."""
    alvo = regex(chave)
    for linha in texto.splitlines():
        m = alvo.search(linha)
        if not m:
            continue
        valor = RE_MONTANTE.match(linha[m.end() :].lstrip())
        if valor:
            return montante(valor.group(0))
    return None
