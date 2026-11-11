"""F8: lista gerada das fontes fiscais (docs/fontes-fiscais.md) e onde cada uma é usada."""

import csv
import re

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


# anotações de desenvolvimento que não podem chegar ao que o usuário lê (site, telas)
INTERNO = re.compile(
    r"(?i)decis(ão|ões)(?: d[ao]s?)?(?: F\d+ e)? do usuário|parecer|especialista|"
    r"(seguido|conferido|escolhida|confirmad[oa]) pelo usuário|mantenedor|\bF\d+\b"
)


def test_publico_tira_as_anotacoes_de_desenvolvimento() -> None:
    casos = {
        "custo mantido (TRF1; decisão do usuário 19/10/2026). A RFB diverge":
            "custo mantido (TRF1; critério do COIR). A RFB diverge",
        "art. 171 §4º). Parecer do especialista de 09/11/2026, seguido pelo usuário.":
            "art. 171 §4º). Critério do COIR.",
        "Day trade pela marcação D (F9, decisão do usuário de 25/10/2026). IN RFB":
            "Day trade pela marcação D. IN RFB",
        "Decisões da F4 (motor fiscal, 14-15/10/2026)": "Critérios do COIR",
        "Casos do especialista fiscal (pareceres de 20 e 21/10/2026)":
            "Casos conferidos à mão (critérios do COIR)",
        "Manual MIR 2026; decisão do usuário (14/10/2026)": "Manual MIR 2026; critério do COIR",
        # datas que citam versão de página oficial ou vigência ficam
        "a página Retenções da RFB (19/11/2025) restringe o DT":
            "a página Retenções da RFB (19/11/2025) restringe o DT",
        "vigência (01/01/2026): alíquota": "vigência (01/01/2026): alíquota",
        "veto derrubado só no DOU de 11/06/2021)": "veto derrubado só no DOU de 11/06/2021)",
    }  # fmt: skip
    for bruto, limpo in casos.items():
        assert ff._publico(bruto) == limpo  # pyright: ignore[reportPrivateUsage]


def test_documento_publicado_sem_anotacoes_internas() -> None:
    for linha in ff.DESTINO.read_text(encoding="utf-8").splitlines():
        assert not INTERNO.search(linha), linha


def test_fonte_dos_parametros_mostrada_em_cadastros_sem_anotacoes_internas() -> None:
    with (ff.RAIZ / "data" / "parametros_fiscais.csv").open(encoding="utf-8") as f:
        for r in csv.DictReader(f):
            assert not INTERNO.search(r["fonte"]), r["nome"]
