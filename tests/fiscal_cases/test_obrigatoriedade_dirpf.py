"""Aviso de obrigatoriedade da declaração pelos dados de bolsa (decisão do usuário de 10/11/2026,
leitura conservadora; IN RFB anuais da DIRPF art. 2º, ficha docs/normas/in-rfb-dirpf-anuais.md):
até o exercício 2022, qualquer operação em bolsa obriga; desde 2023, vendas acima de R$ 40 mil no
ano ou ganho líquido sujeito ao imposto."""

from datetime import date

from coir.engine.obrigatoriedade import aviso_obrigatoriedade
from tests.fiscal_cases.apoio import REFS, Livro


def _aviso(livro: Livro, ano: int) -> str | None:
    ap = livro.apurar(date(ano, 12, 31))
    return aviso_obrigatoriedade(1, ano, ap.movimentos, ap.meses, REFS.parametros)


def test_ate_o_exercicio_2022_qualquer_operacao_obriga() -> None:
    """Só uma compra de 1.000,00 em 2021: declaração de 2022 obrigatória."""
    livro = Livro()
    livro.nota(date(2021, 3, 10), livro.op("C", "PETR4", 100, "10"))
    aviso = _aviso(livro, 2021)
    assert aviso is not None and aviso.startswith("Declaração do exercício 2022 obrigatória")


def test_desde_o_exercicio_2023_so_vendas_acima_de_40_mil_ou_ganho_tributavel() -> None:
    """2024: compra de 30.000 e venda de 3.000 com ganho isento: não obriga. 2025: vendas de
    45.000 em 3 meses, todas isentas: obriga pelas vendas. 2026: uma venda de 25.000 com ganho
    tributável (C 900 VALE3 a 10,00 em janeiro): obriga pelo ganho."""
    livro = Livro()
    livro.nota(date(2024, 2, 5), livro.op("C", "PETR4", 3000, "10"))
    livro.nota(date(2024, 5, 6), livro.op("V", "PETR4", 300, "10"))
    assert _aviso(livro, 2024) is None
    for mes in (3, 4, 5):
        livro.nota(date(2025, mes, 10), livro.op("V", "PETR4", 900, "16.67"))
    aviso = _aviso(livro, 2025)
    assert aviso is not None
    assert "vendas em bolsa de R$ 45.009,00 no ano, acima de R$ 40.000,00" in aviso
    assert "ganho" not in aviso.split(":")[1].split("(")[0]
    livro.nota(date(2026, 1, 12), livro.op("C", "VALE3", 900, "10"))
    livro.nota(date(2026, 3, 10), livro.op("V", "VALE3", 900, "27.78"))
    aviso = _aviso(livro, 2026)
    assert aviso is not None and "ganho líquido sujeito ao imposto" in aviso


def test_ir_minimo_informa_os_ganhos_de_bolsa_desde_2026() -> None:
    """Lei 9.250 art. 16-A §1º I (Lei 15.270/2025), decisão do usuário de 10/11/2026: total
    informativo. 2026: C 1.000 PETR4 a 10,00; V 500 a 20,00 em março (10.000, isento 5.000) e V
    500 a 50,00 em maio (25.000, tributado 20.000, imposto 3.000). Base: 25.000,00. Em 2025,
    nada."""
    from coir.engine.obrigatoriedade import aviso_irpf_minimo

    livro = Livro()
    livro.nota(date(2025, 2, 5), livro.op("C", "VALE3", 100, "10"))
    livro.nota(date(2025, 6, 5), livro.op("V", "VALE3", 100, "12"))
    livro.nota(date(2026, 1, 12), livro.op("C", "PETR4", 1000, "10"))
    livro.nota(date(2026, 3, 10), livro.op("V", "PETR4", 500, "20"))
    livro.nota(date(2026, 5, 11), livro.op("V", "PETR4", 500, "50"))
    ap = livro.apurar(date(2026, 12, 31))
    assert aviso_irpf_minimo(1, 2025, ap.meses, REFS.parametros) is None
    aviso = aviso_irpf_minimo(1, 2026, ap.meses, REFS.parametros)
    assert aviso is not None
    assert "só para quem recebe mais de R$ 600.000,00 no ano" in aviso
    assert "R$ 25.000,00 (tributados R$ 20.000,00 e isentos R$ 5.000,00)" in aviso
    assert "o imposto apurado sobre eles, R$ 3.000,00, é dedutível" in aviso
