from decimal import Decimal

import pytest

from coir.erros import GabaritoError
from coir.gabarito import piloto


def test_interpretar_nome() -> None:
    assert piloto.interpretar_nome("Agora_NC_2008_04.pdf") == piloto.ArquivoNota(
        "AGORA", 2008, 4, False
    )
    assert piloto.interpretar_nome("XP_NC_2023_06_BMF.pdf") == piloto.ArquivoNota(
        "XP", 2023, 6, True
    )
    with pytest.raises(GabaritoError):
        piloto.interpretar_nome("BTG_NC_2023_11_Opcao-Flexível.pdf")


def test_paginas_da_nota_aceita_pontos_e_nao_casa_dentro_de_outro_numero() -> None:
    paginas = ["Nr. nota\n599.721\nFolha 1", "Nr. nota 1599721", "nota 599721 folha"]
    assert piloto.paginas_da_nota(paginas, "599721") == [0, 2]


def test_liquido_pdf() -> None:
    assert piloto.liquido_pdf("Líquido para 24/04/2008 18.686,92 D") == Decimal("-18686.92")
    assert piloto.liquido_pdf("Líquido para 14/08/2026 19.663,11 C") == Decimal("19663.11")
    assert piloto.liquido_pdf("sem resumo") is None


@pytest.mark.parametrize(
    ("liquido", "estimado", "irrf", "esperado"),
    [
        ("100.00", "100.00", "0", "confere"),
        ("100.00", "100.00", "0.98", "confere (IRRF não descontado no líquido)"),
        ("98.64", "100.00", "1.36", "confere (IRRF descontado no líquido)"),
        ("90.00", "100.00", "1.36", "**diferença -10,00**"),
    ],
)
def test_conferir_liquido(liquido: str, estimado: str, irrf: str, esperado: str) -> None:
    assert piloto.conferir_liquido(Decimal(liquido), Decimal(estimado), Decimal(irrf)) == esperado


def test_linhas_resumo_filtra_dados_pessoais() -> None:
    texto = "\n".join(
        [
            "Cliente FULANO DE TAL",
            "C.P.F./C.N.P.J/C.V.M./C.O.B. 000.000.000-00",
            "Valor das operações 18.660,00",
            "0,00 0,00 0,00 0,00 44,00 | C",
            "Rua das Flores, 10",
        ]
    )
    assert piloto.linhas_resumo(texto) == [
        "Valor das operações 18.660,00",
        "0,00 0,00 0,00 0,00 44,00 | C",
    ]
