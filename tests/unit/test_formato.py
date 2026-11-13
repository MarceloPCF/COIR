"""Valores em reais nos textos que o usuário lê saem no padrão brasileiro ("R$ 112.626,22")."""

import ast
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from coir.domain.formato import reais
from coir.engine.obrigatoriedade import aviso_obrigatoriedade
from tests.fiscal_cases.apoio import REFS, Livro

RAIZ = Path(__file__).resolve().parents[2] / "coir"
CAMADAS = ("domain", "engine", "parsing")  # onde nascem os avisos do motor e da leitura das notas


@pytest.mark.parametrize(
    ("valor", "texto"),
    [
        ("112626.22", "R$ 112.626,22"),
        ("40000", "R$ 40.000,00"),
        ("600000.00", "R$ 600.000,00"),
        ("1234567.8", "R$ 1.234.567,80"),
        ("999.99", "R$ 999,99"),
        ("0.05", "R$ 0,05"),
        ("0", "R$ 0,00"),
        ("-1500.5", "R$ -1.500,50"),
    ],
)
def test_reais_no_padrao_brasileiro(valor: str, texto: str) -> None:
    assert reais(Decimal(valor)) == texto


def _reais_crus(caminho: Path) -> list[str]:
    """F-strings com "R$ " seguido de um valor que não passa por ``reais()``."""
    achados: list[str] = []
    for no in ast.walk(ast.parse(caminho.read_text(encoding="utf-8"))):
        if not isinstance(no, ast.JoinedStr):
            continue
        for antes, valor in zip(no.values, no.values[1:], strict=False):
            if not (isinstance(antes, ast.Constant) and str(antes.value).endswith("R$ ")):
                continue
            chamada = valor.value if isinstance(valor, ast.FormattedValue) else None
            if not (isinstance(chamada, ast.Call) and isinstance(chamada.func, ast.Name)
                    and chamada.func.id == "reais"):  # fmt: skip
                achados.append(f"{caminho.relative_to(RAIZ.parent)}:{no.lineno}")
    return achados


def test_avisos_nao_formatam_reais_a_mao() -> None:
    erros = [
        achado
        for camada in CAMADAS
        for caminho in sorted((RAIZ / camada).rglob("*.py"))
        for achado in _reais_crus(caminho)
    ]
    assert not erros, "use coir.domain.formato.reais(): " + ", ".join(erros)


def test_aviso_de_obrigatoriedade_no_padrao_brasileiro() -> None:
    """O caso que apareceu na tela Declaração anual: vendas de 112.626,22 no ano."""
    livro = Livro()
    livro.nota(date(2025, 3, 10), livro.op("C", "PETR4", 10000, "10"))
    livro.nota(date(2025, 6, 10), livro.op("V", "PETR4", 10000, "11.262622"))
    ap = livro.apurar(date(2025, 12, 31))
    aviso = aviso_obrigatoriedade(1, 2025, ap.movimentos, ap.meses, REFS.parametros)
    assert aviso is not None
    assert "vendas em bolsa de R$ 112.626,22 no ano, acima de R$ 40.000,00" in aviso
