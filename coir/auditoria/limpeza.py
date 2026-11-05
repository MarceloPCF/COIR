"""Limpeza dos valores do registro de auditoria, classe por classe.

Vale para a gravação (``limpar``) e para a exportação (``valido``): o portão de privacidade exige
que a limpeza não mude nada no que leu; assim, um arquivo adulterado ou escrito por um programa com
defeito não passa.

Texto livre (mensagem de erro, linha de nota fora do padrão) é o caso difícil: pode trazer nome,
endereço ou o nome de um arquivo. Por isso a regra é de permissão: só ficam as palavras que o
próprio programa escreve (``vocabulario``) e os códigos de negociação. O resto vira a forma:
palavra desconhecida vira ``«L7»`` (7 letras), número vira ``9.999,99``. A forma não identifica
ninguém e mostra o layout, que é o que se precisa para corrigir o leitor.
"""

import re
import unicodedata
from collections.abc import Iterable
from datetime import date, datetime
from decimal import ROUND_HALF_EVEN, Decimal, InvalidOperation

from coir.auditoria.catalogo import PREFIXOS, Classe
from coir.auditoria.vocabulario import PALAVRAS

LIMITE_ROTULO = 60
LIMITE_TEXTO = 600
LIMITE_LINHA = 200  # uma linha do esqueleto de uma página, ou um motivo
LIMITE_LINHAS = 90  # linhas do esqueleto
LIMITE_LISTA = 50
LIMITE_INTEIRO = 10**10  # 11 dígitos já teria cara de CPF
LIMITE_DIFERENCA = Decimal(100)
FIM_CORTADO = " …"

RE_TOKEN = re.compile(
    r"""
      (?P<mascara>«(?:[LA]\d+|sql|caminho|url|email)»)
    | (?P<sql>\[(?:parameters|SQL):.*?\])
    | (?P<caminho>(?:[A-Za-z]:[\\/]|\\\\)[^\s'"<>|]*|(?<![\w.])/(?:[\w.\-]+/)+[\w.\-]*)
    | (?P<url>https?://\S+)
    | (?P<email>[\w.+\-]+@[\w\-]+(?:\.[\w\-]+)+)
    | (?P<numero>\d+(?:[.,/:\-]\d+)*(?![^\W_]))
    | (?P<alfanumerico>[^\W_]+)
    """,
    re.VERBOSE | re.DOTALL,
)
RE_CONTROLE = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\u2028\u2029]")
RE_ESPACOS = re.compile(r"\s+")
# código de negociação: raiz de 4 (a última é letra) e 1 ou 2 dígitos (PETR4, BOVA11, B3SA3, A1AP34,
# WINZ24; F do fracionário), ou raiz, letra da série e até 3 dígitos (PETRK300E, PETRK300W1). Conta
# e CPF com uma letra na frente não têm essa forma: nunca há 4 dígitos em sequência
RE_ATIVO = re.compile(r"[A-Z0-9]{3}[A-Z]\d{1,2}F?|[A-Z0-9]{3}[A-Z][A-X]\d{1,3}(?:E|W\d)?")
# nome de ativo impresso sempre traz a espécie (VALE ON NM, ITAU PN N1); nome de pessoa e endereço
# não trazem, e viram texto limpo
ESPECIES = frozenset(
    [
        "ON",
        "OR",
        "OB",
        "PN",
        "PR",
        "PNA",
        "PNB",
        "PNC",
        "PND",
        "PRA",
        "PRB",
        "UNT",
        "UNIT",
        "CI",
        "ER",
        "DR1",
        "DR2",
        "DR3",
        "DRN",
        "BDR",
        "ETF",
        "FII",
        "FIA",
        "CDA",
        "REC",
        "RCB",
        "DIR",
    ]
)
RE_PSEUDONIMO = re.compile(rf"(?:{'|'.join(PREFIXOS.values())})_[0-9a-f]{{10}}")
RE_DATA = re.compile(r"\d{4}-\d{2}-\d{2}")
RE_MES = re.compile(r"\d{4}-\d{2}")
RE_GRANDEZA = re.compile(r"0|[+-]1e\d{1,2}")
RE_DIFERENCA = re.compile(r"[+-]\d{1,3}\.\d{2}")
RE_IMPRESSO = re.compile(r"[A-Z0-9 ./&\-]{1,40}")
RE_NUMERO_LONGO = re.compile(r"\d+(?:[.\-/]\d+)+|\d{3,}")
RE_CODIGO = re.compile(r"[A-Za-z_<][\w.<>\-]{0,99}")
RE_QUADRO = re.compile(r"[A-Za-z_][\w.\-]*:[\w.<>]+:\d+|[A-Za-z_]\w*")
RE_VERSAO = re.compile(r"[\w.+\-]{1,60}")
RE_ROTA = re.compile(r"/[\w/\-]{0,60}")


def sem_acento(texto: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c))


def normalizada(palavra: str) -> str:
    return sem_acento(palavra).casefold()


def _forma_numero(numero: str) -> str:
    if len(numero) <= 2 and numero.isdigit():
        return numero  # "p. 3", "folha 2": números assim não identificam
    return re.sub(r"\d", "9", numero)


def _token(m: re.Match[str]) -> str:
    tipo, valor = m.lastgroup, m.group()
    if tipo == "mascara":
        return valor
    if tipo in ("sql", "caminho", "url", "email"):
        return f"«{tipo}»"
    if tipo == "numero" or valor.isdigit():
        return _forma_numero(valor)
    if RE_ATIVO.fullmatch(valor):
        return valor
    if valor.casefold() == "b3":  # a única palavra do programa com dígito
        return valor
    if any(c.isdigit() for c in valor):
        return f"«A{len(valor)}»"
    if len(valor) == 1 or normalizada(valor) in PALAVRAS:
        return valor
    return f"«L{len(valor)}»"


def _cortar(texto: str, limite: int) -> str:
    if len(texto) <= limite:
        return texto
    # não corta no meio de um token: o texto cortado tem de continuar limpo
    cabeca = texto[: limite - len(FIM_CORTADO)].rpartition(" ")[0]
    return cabeca + FIM_CORTADO if cabeca.strip() else FIM_CORTADO.strip()


def texto_limpo(texto: str, limite: int = LIMITE_TEXTO, espacos: bool = False) -> str:
    """Só as palavras do vocabulário e os códigos de negociação ficam. ``espacos``: mantém os
    espaços como estão (a linha de uma nota, onde a coluna importa)."""
    limpo = RE_TOKEN.sub(_token, RE_CONTROLE.sub(" ", texto))
    return _cortar(limpo if espacos else RE_ESPACOS.sub(" ", limpo).strip(), limite)


def esqueleto(texto: str) -> tuple[str, ...]:
    """A forma de uma página de nota: o que se precisa ver para ajustar o leitor, sem dado."""
    return tuple(
        limpa
        for linha in texto.splitlines()[:LIMITE_LINHAS]
        if (limpa := texto_limpo(linha.rstrip(), LIMITE_LINHA, espacos=True))
    )


def grandeza(valor: object) -> str:
    """Sinal e ordem de grandeza, nunca o valor (1.234,56 vira +1e3)."""
    numero = _decimal(valor)
    if not numero:
        return "0"
    return f"{'-' if numero < 0 else '+'}1e{max(numero.adjusted(), 0)}"


def diferenca(valor: object) -> str:
    """Até 100 é o que explica a conciliação e não diz nada do cliente; acima, a grandeza."""
    numero = _decimal(valor)
    if abs(numero) > LIMITE_DIFERENCA:
        return grandeza(numero)
    return f"{numero.quantize(Decimal('0.01'), rounding=ROUND_HALF_EVEN):+.2f}"


def _decimal(valor: object) -> Decimal:
    if isinstance(valor, bool) or not isinstance(valor, Decimal | int | float | str):
        raise ValueError("valor sem número")
    try:
        numero = Decimal(str(valor))
    except InvalidOperation as erro:
        raise ValueError("valor sem número") from erro
    if not numero.is_finite():
        raise ValueError("valor sem número")
    return numero


def impresso(valor: str) -> str:
    """Nome de ativo como a nota imprime: é público e é o que explica "ativo sem cadastro", mas
    está fora do vocabulário. Só em maiúsculas, com a espécie (ON, PN...) e sem número longo; outra
    forma vira texto limpo. Se a leitura errar a coluna, o que vier aqui ainda passa pelos termos
    do titular, no portão."""
    limpo = RE_ESPACOS.sub(" ", sem_acento(valor)).strip()[:40].rstrip()
    limpo = RE_NUMERO_LONGO.sub(lambda m: re.sub(r"\d", "9", m.group()), limpo)
    if RE_IMPRESSO.fullmatch(limpo) and ESPECIES & set(limpo.split()):
        return limpo
    return texto_limpo(valor, LIMITE_ROTULO)


def _iso(valor: object, tamanho: int) -> str:
    if isinstance(valor, datetime | date):
        return valor.isoformat()[:tamanho]
    raise ValueError("data esperada")


def _lista(valor: object, limite: int = LIMITE_LISTA) -> list[object]:
    if not isinstance(valor, Iterable) or isinstance(valor, str | bytes):
        raise ValueError("lista esperada")
    return list(valor)[:limite]


def _inteiro(valor: object) -> int:
    if not valido(Classe.INTEIRO, valor):
        raise ValueError("inteiro fora do limite")
    return valor  # type: ignore[return-value]


def _ativo(valor: object) -> str:
    if not valido(Classe.ATIVO, valor):
        raise ValueError("fora da forma de código de negociação")
    return valor  # type: ignore[return-value]


def _na_forma(classe: Classe, valor: object) -> object:
    if not valido(classe, valor):
        raise ValueError(f"fora da forma de {classe.value}")
    return valor


def limpar(classe: Classe, valor: object) -> object:
    """O valor como vai para o arquivo. Levanta ``ValueError`` se ele não couber na classe."""
    match classe:
        case Classe.DATA:
            return _iso(valor, 10)
        case Classe.MES:
            return _iso(valor, 7)
        case Classe.INTEIRO:
            return _inteiro(valor)
        case Classe.INTEIROS:
            return [_inteiro(v) for v in _lista(valor)]
        case Classe.LOGICO:
            return bool(valor)
        case Classe.GRANDEZA:
            return grandeza(valor)
        case Classe.DIFERENCA:
            return diferenca(valor)
        case Classe.ATIVO:
            return _ativo(valor)
        case Classe.ATIVOS:
            return [_ativo(v) for v in _lista(valor)]
        case Classe.ROTULO:
            return texto_limpo(str(valor), LIMITE_ROTULO)
        case Classe.TEXTO:
            return texto_limpo(str(valor))
        case Classe.TEXTOS:
            return [texto_limpo(str(v), LIMITE_LINHA, espacos=True)
                    for v in _lista(valor, LIMITE_LINHAS)]  # fmt: skip
        case Classe.IMPRESSO:
            return impresso(str(valor))
        case Classe.PILHA:
            return _na_forma(classe, [str(v) for v in _lista(valor)])
        case Classe.CODIGO | Classe.VERSAO | Classe.ROTA | Classe.PSEUDONIMO:
            return _na_forma(classe, str(valor))


def _so_textos(valor: object, limite: int) -> list[str] | None:
    """A lista de textos, ou ``None`` se não for só isso."""
    if not isinstance(valor, list) or len(valor) > limite:
        return None
    textos = [v for v in valor if isinstance(v, str)]
    return textos if len(textos) == len(valor) else None


def _data_valida(valor: object, forma: re.Pattern[str], sufixo: str = "") -> bool:
    if not isinstance(valor, str) or not forma.fullmatch(valor):
        return False
    try:
        date.fromisoformat(valor + sufixo)
    except ValueError:
        return False
    return True


def _casa(padrao: re.Pattern[str], valor: object) -> bool:
    return isinstance(valor, str) and padrao.fullmatch(valor) is not None


def _todos(classe: Classe, valor: object) -> bool:
    return isinstance(valor, list) and len(valor) <= LIMITE_LISTA and all(
        valido(classe, v) for v in valor)  # fmt: skip


def valido(classe: Classe, valor: object) -> bool:
    """O valor está na forma exata da classe (a limpeza não mudaria nada)."""
    match classe:
        case Classe.PSEUDONIMO:
            return _casa(RE_PSEUDONIMO, valor)
        case Classe.DATA:
            return _data_valida(valor, RE_DATA)
        case Classe.MES:
            return _data_valida(valor, RE_MES, "-01")
        case Classe.INTEIRO:
            return isinstance(valor, int) and not isinstance(valor, bool) and (
                abs(valor) < LIMITE_INTEIRO)  # fmt: skip
        case Classe.INTEIROS:
            return _todos(Classe.INTEIRO, valor)
        case Classe.LOGICO:
            return isinstance(valor, bool)
        case Classe.GRANDEZA:
            return _casa(RE_GRANDEZA, valor)
        case Classe.DIFERENCA:
            return _casa(RE_DIFERENCA, valor) or _casa(RE_GRANDEZA, valor)
        case Classe.ATIVO:
            return _casa(RE_ATIVO, valor)
        case Classe.ATIVOS:
            return _todos(Classe.ATIVO, valor)
        case Classe.ROTULO:
            return isinstance(valor, str) and texto_limpo(valor, LIMITE_ROTULO) == valor
        case Classe.TEXTO:
            return isinstance(valor, str) and texto_limpo(valor) == valor
        case Classe.TEXTOS:
            textos = _so_textos(valor, LIMITE_LINHAS)
            return textos is not None and all(
                texto_limpo(t, LIMITE_LINHA, espacos=True) == t for t in textos
            )
        case Classe.IMPRESSO:
            return isinstance(valor, str) and impresso(valor) == valor
        case Classe.CODIGO:
            return _casa(RE_CODIGO, valor)
        case Classe.PILHA:
            textos = _so_textos(valor, LIMITE_LISTA)
            return textos is not None and all(RE_QUADRO.fullmatch(t) for t in textos)
        case Classe.VERSAO:
            return _casa(RE_VERSAO, valor)
        case Classe.ROTA:
            return _casa(RE_ROTA, valor)
