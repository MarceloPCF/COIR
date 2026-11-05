"""Portão de privacidade da exportação do registro de auditoria.

Não confia no arquivo: ele pode ter sido escrito por um programa com defeito ou por uma versão
antiga, ou editado à mão. Duas passagens:

1. Linha a linha (``conferir_linhas``): tem de ser um evento do catálogo, só com campos permitidos,
   cada valor na forma exata da sua classe (a limpeza não mudaria nada), sem nenhum formato ou termo
   pessoal. A linha que não passa é retida: fica só a contagem por motivo, nunca o conteúdo.
2. No texto final inteiro (``exigir_texto_limpo``): os formatos e os termos de novo. Aqui qualquer
   achado recusa a exportação toda; não deveria acontecer: é a rede para um defeito na primeira
   passagem.

Os motivos dizem o tipo do achado, nunca o valor.
"""

import json
import re
from collections import Counter
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field
from datetime import datetime
from itertools import pairwise

from coir.auditoria.catalogo import CAMPOS, EVENTOS, Classe
from coir.auditoria.limpeza import normalizada, valido
from coir.auditoria.vocabulario import PALAVRAS
from coir.domain.cpf import so_digitos
from coir.erros import AnonimizacaoError

RE_HORA = re.compile(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}[+-]\d{2}:\d{2}")
RE_SESSAO = re.compile(r"[0-9a-f]{8}")
RE_OP = re.compile(r"[0-9a-f]{6}")
NIVEIS = frozenset("IAEC")
CHAVES = frozenset({"t", "s", "n", "nv", "ev", "op", "d"})
OBRIGATORIAS = CHAVES - {"op"}
# as classes onde um nome caberia; pseudônimo, data e número não têm onde escondê-lo
CLASSES_COM_TEXTO = frozenset(
    {
        Classe.ROTULO, Classe.TEXTO, Classe.TEXTOS, Classe.IMPRESSO, Classe.ATIVO, Classe.ATIVOS,
        Classe.CODIGO, Classe.VERSAO, Classe.ROTA, Classe.PILHA,
    }
)  # fmt: skip

RE_PALAVRA = re.compile(r"[^\W\d_]+")
RE_NUMERO = re.compile(r"\d[\d.\-/ ]*\d")
MINIMO_PALAVRA = 4
MINIMO_NUMERO = 4
PALAVRAS_GENERICAS = frozenset({
    "rua", "avenida", "alameda", "travessa", "praca", "rodovia", "estrada", "largo", "apto",
    "apartamento", "bloco", "casa", "sala", "andar", "conjunto", "edificio", "torre",
    "condominio", "lote", "quadra", "fundos", "loja", "numero", "bairro", "centro", "jardim",
    "vila", "parque", "cidade", "estado", "brasil", "telefone", "celular", "email", "agencia",
    "conta", "corrente", "banco", "assessor",
})  # fmt: skip

# formatos pessoais: uma sequência só de noves é a forma de um número já limpo, não um número
FORMATOS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("CPF", re.compile(r"(?<![\d.])\d{3}[.\s]?\d{3}[.\s]?\d{3}[-.\s]?\d{2}(?!\d)")),
    ("CNPJ", re.compile(r"(?<![\d.])\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}(?!\d)")),
    ("telefone", re.compile(r"(?<![\d.,/-])(?:\(\d{2,3}\)\s?)?\d{4,5}-\d{4}(?![\d-])(?![.,]\d)")),
    ("CEP", re.compile(r"(?<![\d.,/-])(?:\d{5}|\d{2}\.\d{3})-\d{3}(?![\d-])(?![.,]\d)")),
)
RE_EMAIL = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
RE_CAMINHO = re.compile(r"[A-Za-z]:\\{1,2}|\\{1,2}Users\\|/home/|/Users/", re.IGNORECASE)


@dataclass(frozen=True)
class Termos:
    """O que identifica o titular, já na forma de busca. Fica em memória, nunca vai ao arquivo."""

    palavras: frozenset[str] = frozenset()  # sem caixa e sem acento
    # também são palavras do programa: num texto do programa não se distinguem do nome, mas na
    # forma de uma página de nota, onde o nome impresso cairia, não há por que ficarem
    comuns: frozenset[str] = frozenset()
    numeros: tuple[re.Pattern[str], ...] = ()
    expressoes: frozenset[str] = frozenset()  # nome inteiro, sem espaços
    sensiveis: tuple[str, ...] = ()  # lista local (termos-sensiveis.txt), sem caixa
    segredos: tuple[str, ...] = ()

    def __repr__(self) -> str:
        return f"Termos({len(self.palavras)} palavra(s), {len(self.numeros)} número(s))"

    @classmethod
    def de(
        cls, valores: Iterable[str], sensiveis: Sequence[str] = (), segredos: Sequence[str] = ()
    ) -> "Termos":
        """``valores``: nome, CPF, contas, endereço... Palavra que o próprio programa usa não serve
        de termo: não se distingue dela no texto, e o resto do nome já não passa pela limpeza; o
        nome inteiro é procurado como expressão."""
        palavras: set[str] = set()
        comuns: set[str] = set()
        numeros: set[str] = set()
        expressoes: set[str] = set()
        for valor in valores:
            partes = [normalizada(p) for p in RE_PALAVRA.findall(valor)]
            for n in partes:
                if len(n) >= MINIMO_PALAVRA and n not in PALAVRAS_GENERICAS:
                    (comuns if n in PALAVRAS else palavras).add(n)
            if not any(c.isdigit() for c in valor):  # duas palavras seguidas do nome já identificam
                expressoes.update(a + b for a, b in pairwise(partes))
            for m in RE_NUMERO.findall(valor):
                d = so_digitos(m).lstrip("0")
                if len(d) >= MINIMO_NUMERO:
                    numeros.add(d)
            if " " in valor.strip() and not any(c.isdigit() for c in valor):
                expressoes.add(_compacta(valor))
        padroes = tuple(re.compile(r"(?<!\d)" + r"[.\- ]?".join(n) + r"(?!\d)")
                        for n in sorted(numeros))  # fmt: skip
        return cls(frozenset(palavras), frozenset(comuns), padroes, frozenset(expressoes),
                   tuple(s.casefold() for s in sensiveis if s),
                   tuple(s for s in segredos if s))  # fmt: skip


def _compacta(texto: str) -> str:
    return "".join(RE_PALAVRA.findall(normalizada(texto)))


def _termo_no_texto(texto: str, termos: Termos) -> str | None:
    if set(RE_PALAVRA.findall(normalizada(texto))) & termos.palavras:
        return "palavra do titular"
    if any(p.search(texto) for p in termos.numeros):
        return "número do titular"
    compacto = _compacta(texto)
    if any(e in compacto for e in termos.expressoes):
        return "nome do titular"
    caixa = texto.casefold()
    if any(s in caixa for s in termos.sensiveis):
        return "termo da lista local"
    if any(s in texto for s in termos.segredos):
        return "segredo do registro"
    return None


def _formato_pessoal(texto: str) -> str | None:
    for nome, padrao in FORMATOS:
        for m in padrao.finditer(texto):
            # o DDD (até 2 dígitos) fica na limpeza: a forma é o número que vem depois
            if set(so_digitos(m.group().rpartition(")")[2])) != {"9"}:
                return nome
    if RE_EMAIL.search(texto):
        return "e-mail"
    if RE_CAMINHO.search(texto):
        return "caminho de pasta"
    return None


def _textos_do_campo(classe: Classe, valor: object) -> list[str]:
    if classe not in CLASSES_COM_TEXTO:
        return []
    if isinstance(valor, list):
        textos = [v for v in valor if isinstance(v, str)]
        # o número da linha da pilha não é um número do titular
        return [t.rsplit(":", 1)[0] for t in textos] if classe is Classe.PILHA else textos
    return [valor] if isinstance(valor, str) else []


def _termo_comum(classe: Classe, texto: str, termos: Termos) -> str | None:
    if classe not in (Classe.TEXTOS, Classe.IMPRESSO):
        return None
    achou = set(RE_PALAVRA.findall(normalizada(texto))) & termos.comuns
    return "palavra do titular" if achou else None


def conferir_linha(linha: str, termos: Termos) -> tuple[dict[str, object] | None, str]:
    """O evento conferido ou, se não passar, ``None`` e o motivo (nunca o valor)."""
    try:
        objeto = json.loads(linha)
    except ValueError:
        return None, "linha ilegível"
    if not isinstance(objeto, dict):
        return None, "linha fora do formato"
    if set(objeto) - CHAVES or not set(objeto) >= OBRIGATORIAS:
        return None, "linha fora do formato"
    if motivo := _envelope_invalido(objeto):
        return None, motivo
    nome, dados = objeto["ev"], objeto["d"]
    if not isinstance(nome, str) or nome not in EVENTOS:
        return None, "evento fora do catálogo"
    if not isinstance(dados, dict):
        return None, "linha fora do formato"
    for campo, valor in dados.items():
        if campo not in EVENTOS[nome]:
            return None, "campo fora do catálogo"
        classe = CAMPOS[campo]
        if not valido(classe, valor):
            return None, f"campo fora da forma de {classe.value}"
        for texto in _textos_do_campo(classe, valor):
            achado = (_formato_pessoal(texto) or _termo_no_texto(texto, termos)
                      or _termo_comum(classe, texto, termos))  # fmt: skip
            if achado:
                return None, achado
    return objeto, ""


def _hora_valida(texto: str) -> bool:
    try:
        datetime.fromisoformat(texto)
    except ValueError:
        return False
    return True


def _envelope_invalido(o: dict[str, object]) -> str | None:
    t, s, n, nv, op = o["t"], o["s"], o["n"], o["nv"], o.get("op")
    if not (isinstance(t, str) and RE_HORA.fullmatch(t) and _hora_valida(t)):
        return "hora fora do formato"
    if not (isinstance(s, str) and RE_SESSAO.fullmatch(s)):
        return "sessão fora do formato"
    if not (isinstance(n, int) and not isinstance(n, bool) and 0 <= n < 10**9):
        return "sequência fora do formato"
    if not (isinstance(nv, str) and nv in NIVEIS):
        return "nível fora do formato"
    if op is not None and not (isinstance(op, str) and RE_OP.fullmatch(op)):
        return "operação fora do formato"
    return None


@dataclass
class Conferencia:
    linhas: list[str] = field(default_factory=list[str])  # já na forma canônica
    retidas: Counter[str] = field(default_factory=Counter[str])  # motivo -> quantas


def conferir_linhas(linhas: Iterable[str], termos: Termos) -> Conferencia:
    """Primeira passagem. O texto refeito do que foi conferido é o que vale (sem chave repetida,
    sem espaço a mais)."""
    resultado = Conferencia()
    for linha in linhas:
        objeto, motivo = conferir_linha(linha, termos)
        if objeto is None:
            resultado.retidas[motivo] += 1
        else:
            resultado.linhas.append(json.dumps(objeto, ensure_ascii=False, separators=(",", ":")))
    return resultado


def achados_no_texto(texto: str, termos: Termos) -> list[str]:
    """Segunda passagem, no texto final: os tipos do que achou (vazio é o único aceito)."""
    achados: set[str] = set()
    for linha in texto.splitlines():
        try:
            objeto = json.loads(linha)
            valores = objeto["d"].items()
        except (ValueError, KeyError, AttributeError, TypeError):
            achados.add("linha fora do formato")
            continue
        for campo, valor in valores:
            classe = CAMPOS.get(campo, Classe.TEXTO)
            for texto_campo in _textos_do_campo(classe, valor):
                if achado := (_termo_no_texto(texto_campo, termos)
                              or _termo_comum(classe, texto_campo, termos)):  # fmt: skip
                    achados.add(achado)
        if achado := _formato_pessoal(json.dumps(objeto["d"], ensure_ascii=False)):
            achados.add(achado)
    if any(s in texto for s in termos.segredos):
        achados.add("segredo do registro")
    return sorted(achados)


def exigir_texto_limpo(texto: str, termos: Termos) -> None:
    if achados := achados_no_texto(texto, termos):
        raise AnonimizacaoError(
            "o registro não foi exportado: o portão de privacidade achou " + ", ".join(achados)
        )
