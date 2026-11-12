"""Textos que chegam ao usuário (telas, avisos, relatórios) sem anotações de desenvolvimento.

Quem decidiu, quando e em que fase do projeto fica nas docstrings, nos comentários e em
docs/regras-fiscais.md; nas strings do programa, o critério aparece como "critério do COIR".
O gabarito (comparação com a planilha do mantenedor) é ferramenta do desenvolvedor e fica de fora.
"""

import ast
import re
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
FORA = {"gabarito", "migracoes"}  # ferramentas do desenvolvedor e histórico do banco
INTERNO = re.compile(
    r"(?i)decis(ão|ões)(?: d[ao]s?)?(?: F\d+ e)? do usuário|parecer|especialista|usuário de \d|"
    r"(seguido|conferido|escolhida|confirmad[oa]) pelo usuário|mantenedor|"
    r"\b(?:da|na|pela) F\d+\b"
)


def _docstrings(arvore: ast.AST) -> set[int]:
    ids: set[int] = set()
    for no in ast.walk(arvore):
        if isinstance(no, ast.Module | ast.FunctionDef | ast.AsyncFunctionDef | ast.ClassDef):
            primeiro = no.body[0] if no.body else None
            if isinstance(primeiro, ast.Expr) and isinstance(primeiro.value, ast.Constant):
                ids.add(id(primeiro.value))
    return ids


def _textos(caminho: Path) -> list[tuple[int, str]]:
    arvore = ast.parse(caminho.read_text(encoding="utf-8"))
    docs = _docstrings(arvore)
    achados: list[tuple[int, str]] = []
    for no in ast.walk(arvore):
        if isinstance(no, ast.JoinedStr):  # f-string: as partes fixas
            partes = [v.value for v in no.values if isinstance(v, ast.Constant)]
            achados.append((no.lineno, "".join(str(p) for p in partes)))
        elif isinstance(no, ast.Constant) and isinstance(no.value, str) and id(no) not in docs:
            achados.append((no.lineno, no.value))
    return achados


def test_strings_do_programa_sem_anotacoes_de_desenvolvimento() -> None:
    erros = [
        f"{caminho.relative_to(RAIZ)}:{linha}: {texto[:90]!r}"
        for caminho in sorted((RAIZ / "coir").rglob("*.py"))
        if not FORA & set(caminho.parts)
        for linha, texto in _textos(caminho)
        if INTERNO.search(texto)
    ]
    assert not erros, "\n".join(erros)
