"""Dependências entre camadas e guardas das regras 1 e 6 (CLAUDE.md).

- engine e parsing dependem só de domain (e de coir.erros); domain não depende de outra camada.
- gabarito (ferramenta de validação) pode usar parsing, pyxlsb e openpyxl (planilha externa),
  mas não db, ui nem engine.
- db (SQLAlchemy/Alembic) depende só de domain; servicos orquestra parsing + db (sem ui).
- reports (xlsx/csv/pdf, ADR 0003) parte só do modelo neutro e do domain: sem banco, motor nem UI;
  só ela usa xlsxwriter e reportlab, e não lê o relógio (a data de geração vem de fora).
- Nenhuma dessas camadas importa banco, UI, CLI ou bibliotecas de planilha.
- engine e domain: sem ``float`` (dinheiro é Decimal) e sem leitura de relógio/aleatoriedade.
"""

import ast
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[2] / "coir"
INTERNOS_PERMITIDOS = {
    "domain": ("coir.domain", "coir.erros"),
    "engine": ("coir.engine", "coir.domain", "coir.erros"),
    "parsing": ("coir.parsing", "coir.domain", "coir.erros"),
    # ferramenta de validação: lê a planilha do usuário e PDFs; nunca banco, UI nem motor
    "gabarito": (
        "coir.gabarito", "coir.domain", "coir.erros", "coir.parsing", "coir.recursos",
    ),
    "db": ("coir.db", "coir.domain", "coir.erros"),
    "reports": ("coir.reports", "coir.domain", "coir.erros"),
    "servicos": (
        "coir.servicos", "coir.db", "coir.domain", "coir.engine", "coir.erros", "coir.parsing",
        "coir.reports",
    ),
}  # fmt: skip
EXTERNOS_LIBERADOS = {
    "gabarito": ("pyxlsb", "openpyxl"),
    "db": ("sqlalchemy", "alembic", "sqlite3"),
    "reports": ("xlsxwriter", "reportlab"),
    "servicos": ("sqlalchemy",),
}
EXTERNOS_PROIBIDOS = (
    "sqlalchemy", "alembic", "sqlite3", "nicegui", "webview", "typer",
    "openpyxl", "xlsxwriter", "pyxlsb", "reportlab",
)  # fmt: skip
NAO_DETERMINISTICOS = ("random", "time", "secrets", "uuid")
ATRIBUTOS_RELOGIO = {"now", "today", "utcnow"}


def _modulo_do_arquivo(arquivo: Path) -> list[str]:
    partes = ["coir", *arquivo.relative_to(RAIZ).with_suffix("").parts]
    return partes[:-1] if partes[-1] == "__init__" else partes


def _imports(arquivo: Path, codigo: str | None = None) -> set[str]:
    """Nomes de módulos importados, com imports relativos resolvidos e ``from X import Y``
    gerando ``X`` e ``X.Y``."""
    fonte = arquivo.read_text(encoding="utf-8") if codigo is None else codigo
    pacote = _modulo_do_arquivo(arquivo)
    if arquivo.name != "__init__.py":
        pacote = pacote[:-1]
    nomes: set[str] = set()
    for no in ast.walk(ast.parse(fonte)):
        if isinstance(no, ast.Import):
            nomes.update(a.name for a in no.names)
        elif isinstance(no, ast.ImportFrom):
            if no.level:
                base = pacote[: len(pacote) - (no.level - 1)]
                modulo = ".".join([*base, *([no.module] if no.module else [])])
            else:
                modulo = no.module or ""
            nomes.add(modulo)
            nomes.update(f"{modulo}.{a.name}" for a in no.names)
    return nomes


def _casa(modulo: str, prefixos: tuple[str, ...]) -> bool:
    return any(modulo == p or modulo.startswith(p + ".") for p in prefixos)


def _violacoes_de_import(camada: str, modulos: set[str]) -> list[str]:
    liberados = EXTERNOS_LIBERADOS.get(camada, ())
    proibidos = tuple(p for p in EXTERNOS_PROIBIDOS if p not in liberados) + (
        NAO_DETERMINISTICOS if camada != "parsing" else ()
    )

    def errado(mod: str) -> bool:
        # `from coir.domain import Nota` gera "coir.domain.Nota", que casa o prefixo permitido
        if mod.startswith("coir."):
            return not _casa(mod, INTERNOS_PERMITIDOS[camada])
        return _casa(mod, proibidos)

    return [mod for mod in sorted(modulos) if errado(mod)]


def _violacoes_de_determinismo(codigo: str) -> list[str]:
    erradas: list[str] = []
    for no in ast.walk(ast.parse(codigo)):
        if isinstance(no, ast.Name) and no.id == "float":
            erradas.append(f"linha {no.lineno}: float")
        elif isinstance(no, ast.Constant) and isinstance(no.value, float):
            erradas.append(f"linha {no.lineno}: literal float")
        elif isinstance(no, ast.Attribute) and no.attr in ATRIBUTOS_RELOGIO:
            erradas.append(f"linha {no.lineno}: .{no.attr}")
    return erradas


def _arquivos(camada: str) -> list[Path]:
    arquivos = sorted((RAIZ / camada).rglob("*.py"))
    assert arquivos, f"camada {camada} sem arquivos"
    return arquivos


@pytest.mark.parametrize("camada", sorted(INTERNOS_PERMITIDOS))
def test_camada_respeita_dependencias(camada: str) -> None:
    violacoes = [
        f"{arq.relative_to(RAIZ)}: {mod}"
        for arq in _arquivos(camada)
        for mod in _violacoes_de_import(camada, _imports(arq))
    ]
    assert violacoes == []


@pytest.mark.parametrize("camada", ["engine", "domain"])
def test_camada_sem_float_nem_relogio(camada: str) -> None:
    violacoes = [
        f"{arq.relative_to(RAIZ)} {v}"
        for arq in _arquivos(camada)
        for v in _violacoes_de_determinismo(arq.read_text(encoding="utf-8"))
    ]
    assert violacoes == []


def test_reports_nao_le_o_relogio() -> None:
    """A data de geração entra como parâmetro (relatórios reproduzíveis). Float é permitido aqui:
    medidas de layout do PDF e a gravação de números no xlsx (ADR 0003)."""
    violacoes = [
        f"{arq.relative_to(RAIZ)} {v}"
        for arq in _arquivos("reports")
        for v in _violacoes_de_determinismo(arq.read_text(encoding="utf-8"))
        if "float" not in v
    ]
    assert violacoes == []


@pytest.mark.parametrize(
    ("codigo", "esperado"),
    [
        ("import sqlalchemy.orm", ["sqlalchemy.orm"]),
        ("from coir.db import repo", ["coir.db", "coir.db.repo"]),
        ("from coir import db", ["coir.db"]),
        ("from coir import ui", ["coir.ui"]),
        ("from ..db import x", ["coir.db", "coir.db.x"]),
        ("from .. import db", ["coir.db"]),
        ("from .. import parsing", ["coir.parsing"]),
        ("import sqlite3", ["sqlite3"]),
        ("import random", ["random"]),
        ("from coir.config import dados_dir", ["coir.config", "coir.config.dados_dir"]),
        ("from decimal import Decimal", []),
        ("from coir.domain import Nota", []),
        ("from coir import erros", []),
        ("from .regras import x", []),
    ],
)
def test_detector_de_imports(codigo: str, esperado: list[str]) -> None:
    arquivo = RAIZ / "engine" / "pm.py"  # arquivo hipotético em coir/engine
    assert _violacoes_de_import("engine", _imports(arquivo, codigo)) == esperado


@pytest.mark.parametrize(
    ("codigo", "esperado"),
    [
        ("x: float = 0", ["linha 1: float"]),
        ("y = 0.1", ["linha 1: literal float"]),
        ("from datetime import date\nh = date.today()", ["linha 2: .today"]),
        ("import datetime\nagora = datetime.datetime.now()", ["linha 2: .now"]),
        ("from decimal import Decimal\nv = Decimal('0.1')", []),
    ],
)
def test_detector_de_determinismo(codigo: str, esperado: list[str]) -> None:
    assert _violacoes_de_determinismo(codigo) == esperado
