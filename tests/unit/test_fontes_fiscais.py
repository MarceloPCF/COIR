"""F8: lista gerada das fontes fiscais (docs/fontes-fiscais.md) e onde cada uma é usada."""

import csv

from tools import fontes_fiscais as ff


def test_citacoes_normalizadas() -> None:
    texto = ("Lei 11.033/2004 art. 2º; LC 224/2025 art. 8º; IN RFB 1.585/2015 art. 58; "
             "MP 497/2010; SC Cosit 224/2014; RIR/2018 art. 915; P&R IRPF 2026 pergunta 714; "
             "Res. CMN 4.880/2020; Lei 6.404 art. 252; CTN art. 138; Súmula 360; "
             "REsp 1.149.022/SP; fichas do programa IRPF")  # fmt: skip
    assert ff.normas(texto) == [
        "LC 224/2025", "Lei 11.033/2004", "Lei 6.404", "MP 497/2010", "IN RFB 1.585/2015",
        "SC Cosit 224/2014", "Res. CMN 4.880/2020", "RIR/2018", "P&R IRPF 2026", "CTN",
        "Súmula STJ 360", "REsp 1.149.022/SP", "Programa IRPF",
    ]  # fmt: skip
    assert ff.normas("Lei Complementar 224/2025") == ["LC 224/2025"]


def test_documento_em_dia() -> None:
    """Mudou uma fonte em data/, nas regras ou num caso fiscal: gere o documento de novo
    (uv run python -m tools.fontes_fiscais)."""
    assert ff.DESTINO.read_text(encoding="utf-8") == ff.gerar()


def test_toda_fonte_dos_parametros_esta_no_documento() -> None:
    documento = ff.DESTINO.read_text(encoding="utf-8")
    with (ff.RAIZ / "data" / "parametros_fiscais.csv").open(encoding="utf-8") as f:
        for r in csv.DictReader(f):
            for norma in ff.normas(r["fonte"]):
                assert f"## {norma}" in documento or f"## {norma}/" in documento, norma
            assert f"— {r['nome']}:" in documento or not ff.normas(r["fonte"]), r["nome"]
