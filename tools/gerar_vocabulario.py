"""Gera ``coir/auditoria/vocabulario.py``: as palavras que o texto do registro de auditoria mantém.

São as palavras dos textos do próprio programa (mensagens de erro, rótulos dos layouts das notas e
das planilhas) e nomes das corretoras e instituições. Palavra fora da lista vira só a forma
(``«L7»``), então o arquivo pode estar um pouco defasado sem risco; rode de novo quando as mensagens
mudarem muito (um teste avisa). Docstrings e comentários ficam de fora: só o que o programa escreve.

Sobrenomes e prenomes comuns são tirados mesmo quando também são palavras comuns ("campos"): nesse
texto, perder uma palavra custa menos do que deixar passar um nome.

Uso: ``uv run python tools/gerar_vocabulario.py``
"""

import ast
import csv
import re
import textwrap
import unicodedata
from collections.abc import Iterable, Iterator
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
SAIDA = RAIZ / "coir" / "auditoria" / "vocabulario.py"
RE_PALAVRA = re.compile(r"[^\W\d_]+")
MINIMO = 2
MAXIMO = 30
TABELAS = (("data/instituicoes_b3.csv", "corretora"),)

# mensagens comuns de sistema e de bibliotecas, que não vêm do código do COIR
TECNICAS = """
    access accessed address already and argument attribute bad been cannot column connection could
    database decode denied directory disk does encoding error errno exception exists expected failed
    false file files found full function handle has httperror illegal index integer invalid
    is it key list locked memory missing module must name no none not object of operation out path
    permission process read refused required reset server sqlite such syntax system table the
    timed timeout too true type unable unexpected unicode unknown url value version was with write
    windows python pdf xml json zip http https webview
"""

# prenomes e sobrenomes mais comuns no Brasil: nunca ficam no texto, mesmo sendo também palavras
NOMES_COMUNS = """
    ana antonio francisco carlos paulo pedro lucas luiz luis marcos gabriel rafael daniel marcelo
    bruno eduardo felipe rodrigo manoel mateus andre fernando fabio leonardo gustavo guilherme
    leandro tiago thiago anderson ricardo vinicius sergio jorge alexandre roberto julio cesar
    maria francisca adriana juliana marcia fernanda patricia aline sandra camila amanda bruna
    jessica leticia julia luciana vanessa mariana gabriela vera renata lucia beatriz
    silva santos oliveira souza sousa rodrigues ferreira alves pereira lima gomes ribeiro carvalho
    almeida lopes soares fernandes vieira barbosa rocha dias nascimento andrade moreira nunes
    marques machado mendes freitas cardoso ramos goncalves santana teixeira araujo costa martins
    correia campos cavalcante cavalcanti melo mello pinto moura castro batista
    fulano beltrano sicrano
"""


def _normal(palavra: str) -> str:
    sem = "".join(c for c in unicodedata.normalize("NFKD", palavra) if not unicodedata.combining(c))
    return sem.casefold()


def _palavras(texto: str) -> Iterator[str]:
    for p in RE_PALAVRA.findall(texto):
        n = _normal(p)
        if MINIMO <= len(n) <= MAXIMO:
            yield n


def _docstrings(arvore: ast.Module) -> set[int]:
    """As constantes que são docstrings (pelo id do nó)."""
    nos: set[int] = set()
    for no in ast.walk(arvore):
        if isinstance(no, ast.Module | ast.ClassDef | ast.FunctionDef | ast.AsyncFunctionDef):
            corpo = no.body
            if (
                corpo
                and isinstance(corpo[0], ast.Expr)
                and isinstance(corpo[0].value, ast.Constant)
            ):
                nos.add(id(corpo[0].value))
    return nos


def _do_codigo(arquivo: Path) -> Iterator[str]:
    arvore = ast.parse(arquivo.read_text(encoding="utf-8"))
    docstrings = _docstrings(arvore)
    for no in ast.walk(arvore):
        if isinstance(no, ast.Constant) and isinstance(no.value, str) and id(no) not in docstrings:
            yield from _palavras(no.value)


def _das_tabelas() -> Iterator[str]:
    for relativo, coluna in TABELAS:
        with (RAIZ / relativo).open(encoding="utf-8", newline="") as entrada:
            for linha in csv.DictReader(entrada):
                yield from _palavras(linha[coluna])


def vocabulario() -> set[str]:
    palavras: set[str] = set(_palavras(TECNICAS))
    for arquivo in sorted((RAIZ / "coir").rglob("*.py")):
        if arquivo != SAIDA:
            palavras.update(_do_codigo(arquivo))
    for perfil in sorted((RAIZ / "coir" / "parsing" / "profiles").glob("*.yaml")):
        palavras.update(_palavras(perfil.read_text(encoding="utf-8")))
    palavras.update(_das_tabelas())
    return palavras - set(_palavras(NOMES_COMUNS))


def conteudo(palavras: Iterable[str]) -> str:
    corpo = "\n".join(textwrap.wrap(" ".join(sorted(palavras)), width=99))
    return (
        '"""Gerado por tools/gerar_vocabulario.py; não edite à mão."""\n\n'
        f'_TEXTO = """\n{corpo}\n"""\nPALAVRAS = frozenset(_TEXTO.split())\n'
    )


def main() -> None:
    palavras = vocabulario()
    SAIDA.write_text(conteudo(palavras), encoding="utf-8", newline="\n")
    print(f"{len(palavras)} palavras em {SAIDA.relative_to(RAIZ)}")


if __name__ == "__main__":
    main()
