from decimal import Decimal

from coir.gabarito.conferencia import LinhaNegocio, linhas_bmf, linhas_bovespa, valores_impressos


def test_linhas_bovespa_layouts() -> None:
    texto = "\n".join(
        [
            "BOVESPA C VISTA GERDAU PN N1 H 300 62,20 D 18.660,00",  # Ágora: D/C antes do valor
            "1-BOVESPA V VISTA COPASA ON NM @# 1.300 55,06 71.578,00 C",  # XP
            "1-BOVESPA V OPCAO DE VENDA 12/24 BBDCX512 PN 1000 1,25 1.250,00 C",  # BTG, sem milhar
            "1-BOVESPA V OPCAO DE COMPRA 06/24 VALEF730 ON D 400 1,35 540,00 C",  # obs. D
            "Resumo dos Negócios",
        ]
    )
    assert linhas_bovespa(texto) == [
        LinhaNegocio("C", Decimal(300), Decimal("62.20"), Decimal("18660.00")),
        LinhaNegocio("V", Decimal(1300), Decimal("55.06"), Decimal("71578.00")),
        LinhaNegocio("V", Decimal(1000), Decimal("1.25"), Decimal("1250.00")),
        LinhaNegocio("V", Decimal(400), Decimal("1.35"), Decimal("540.00")),
    ]


def test_linhas_bmf_ignora_ajuste_de_posicao() -> None:
    texto = "\n".join(
        [
            "C WING19 @13/02/2019 1 97.810,0000 DAYTRADE 25,00 C 0,20",
            "V WINM25 18/06/2025 5 137.945,0000 DAY TRADE 387,00 C 0,00",
            "C BGI G22 @25/02/2022 1 343,2500 NORMAL 115,50 D 14,00",
            "V CCM H24 @15/03/2024 6 64,30 NORMAL 0,00 24,00",  # abertura: 0,00 sem D/C
            "V CCMPPPP 0 0,0000 TX. PERMANÊNCIA 0,00 C 0,00",
            "V CCMU23 15/09/2023 2 58,8400 AJUPOS 1.593,00 D 0,00",
        ]
    )
    assert [(x.cv, x.quantidade, x.preco) for x in linhas_bmf(texto)] == [
        ("C", Decimal(1), Decimal("97810.0000")),
        ("V", Decimal(5), Decimal("137945.0000")),
        ("C", Decimal(1), Decimal("343.2500")),
        ("V", Decimal(6), Decimal("64.30")),
    ]


def test_valores_impressos() -> None:
    assert valores_impressos("0,00| 0,41 1,60 | D 1000,00 40,87 | C") == {
        Decimal("0.00"), Decimal("0.41"), Decimal("1.60"), Decimal("1000.00"), Decimal("40.87"),
    }  # fmt: skip
