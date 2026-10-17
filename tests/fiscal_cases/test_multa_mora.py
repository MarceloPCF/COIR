"""Multa de mora da DARF 6015 paga em atraso (números conferidos à mão em 17/10/2026).

Fontes: Lei 9.430/1996 art. 61 caput (0,33% por dia), §1º (contagem a partir do dia seguinte
ao vencimento até o pagamento) e §2º (teto de 20%); RFB, "Como calcular multa de mora" e
glossário do Sicalc (contagem a partir do primeiro dia útil seguinte ao vencimento; pagamento
em dia não útil vale no próximo dia útil); art. 68 §1º (valor abaixo de R$ 10 acumula e vence
com o último período). Decisões do usuário (17/10/2026): contagem como a Receita (``rfb``),
centavos truncados (ROUND_DOWN) e data não útil ajustada.
"""

from datetime import date
from decimal import Decimal

import pytest

from coir.engine.darf import multa_mora
from coir.erros import MotorError
from tests.fiscal_cases.apoio import REFS, parametros

D = Decimal
CAL = REFS.calendario
PADRAO = REFS.parametros


def _multa(principal: str, vencimento: date, pagamento: date, **troca: str) -> tuple[int, D, D]:
    params = parametros(**troca) if troca else PADRAO
    m = multa_mora(D(principal), date(vencimento.year, 1, 1), vencimento, pagamento, params, CAL)
    return m.dias, m.percentual, m.valor


@pytest.mark.parametrize(
    ("caso", "principal", "vencimento", "pagamento", "dias", "percentual", "multa"),
    [
        # C1: pago no vencimento (sexta-feira, último dia útil de outubro): sem multa
        ("C1", "523.47", date(2025, 10, 31), date(2025, 10, 31), 0, "0", "0.00"),
        # C1b: pago antes do vencimento
        ("C1b", "523.47", date(2025, 10, 31), date(2025, 10, 20), 0, "0", "0.00"),
        # C2: um dia, meio de semana: 857,93 x 0,33% = 2,831169 -> 2,83
        ("C2", "857.93", date(2026, 6, 30), date(2026, 7, 1), 1, "0.0033", "2.83"),
        # C3: vence sexta 31/10/2025, paga segunda 03/11: a Receita conta só 03/11 (1 dia);
        # 2.450,00 x 0,33% = 8,085 -> 8,08 truncado
        ("C3", "2450.00", date(2025, 10, 31), date(2025, 11, 3), 1, "0.0033", "8.08"),
        # C4: vence quinta 30/04/2026; 01/05 feriado e fim de semana; paga 04/05: 1 dia
        ("C4", "1500.00", date(2026, 4, 30), date(2026, 5, 4), 1, "0.0033", "4.95"),
        # C5: vence 30/12/2025 (31/12 sem expediente, 01/01 feriado); paga 02/01/2026: 1 dia
        ("C5", "700.00", date(2025, 12, 30), date(2026, 1, 2), 1, "0.0033", "2.31"),
        # C6a: 60 dias (01/08 é útil): 19,80% de 3.000,00 = 594,00
        ("C6a", "3000.00", date(2025, 7, 31), date(2025, 9, 29), 60, "0.1980", "594.00"),
        # C6b: 61 dias = 20,13% -> teto de 20% = 600,00
        ("C6b", "3000.00", date(2025, 7, 31), date(2025, 9, 30), 61, "0.20", "600.00"),
        # C7: atraso longo (104 dias contados de 02/01/2026): teto, 2.000,00
        ("C7", "10000.00", date(2025, 12, 30), date(2026, 4, 15), 104, "0.20", "2000.00"),
        # C8: 7 dias = 2,31% de 1.234,57 = 28,518567 -> 28,51 truncado
        ("C8", "1234.57", date(2025, 9, 30), date(2025, 10, 7), 7, "0.0231", "28.51"),
        # C9: jan/26 6,40 + fev/26 5,10 = 11,50 acumulado vence com fevereiro (31/03/2026);
        # pago em 01/04: 1 dia = 0,03795 -> 0,03 (nada contado desde o vencimento de janeiro)
        ("C9", "11.50", date(2026, 3, 31), date(2026, 4, 1), 1, "0.0033", "0.03"),
    ],
)
def test_multa_de_mora_como_a_receita(
    caso: str, principal: str, vencimento: date, pagamento: date, dias: int, percentual: str,
    multa: str,
) -> None:  # fmt: skip
    """Lei 9.430/1996 art. 61; contagem e centavos como decididos (rfb, ROUND_DOWN)."""
    assert _multa(principal, vencimento, pagamento) == (dias, D(percentual), D(multa)), caso


@pytest.mark.parametrize(
    ("principal", "vencimento", "pagamento", "dias", "multa"),
    [
        ("2450.00", date(2025, 10, 31), date(2025, 11, 3), 3, "24.25"),  # C3: 3 dias corridos
        ("1500.00", date(2026, 4, 30), date(2026, 5, 4), 4, "19.80"),  # C4
        ("700.00", date(2025, 12, 30), date(2026, 1, 2), 3, "6.93"),  # C5
        ("10000.00", date(2025, 12, 30), date(2026, 4, 15), 106, "2000.00"),  # C7
    ],
)
def test_contagem_literal_da_lei_por_parametro(
    principal: str, vencimento: date, pagamento: date, dias: int, multa: str
) -> None:
    """Lei 9.430/1996 art. 61 §1º: do dia seguinte ao vencimento, em dias corridos."""
    resultado = _multa(principal, vencimento, pagamento, multa_mora_inicio_contagem="lei")
    assert (resultado[0], resultado[2]) == (dias, D(multa))


def test_arredondamento_por_parametro() -> None:
    """Lei omissa sobre centavos: 28,518567 -> 28,52 e 8,085 -> 8,09 com ROUND_HALF_UP."""
    troca = {"multa_mora_arredondamento": "ROUND_HALF_UP"}
    assert _multa("1234.57", date(2025, 9, 30), date(2025, 10, 7), **troca)[2] == D("28.52")
    assert _multa("2450.00", date(2025, 10, 31), date(2025, 11, 3), **troca)[2] == D("8.09")


def test_pagamento_em_dia_nao_util_vale_no_proximo_dia_util() -> None:
    """Glossário do Sicalc: arrecadação em sábado, domingo ou feriado vale no próximo dia útil.
    Vence sexta 31/10/2025; informado sábado 01/11 -> segunda 03/11: 1 dia (rfb)."""
    m = multa_mora(D("2450.00"), date(2025, 1, 1), date(2025, 10, 31), date(2025, 11, 1),
                   PADRAO, CAL)  # fmt: skip
    assert (m.pagamento, m.ajustado, m.dias, m.valor) == (date(2025, 11, 3), True, 1, D("8.08"))


def test_parametro_invalido_falha_alto() -> None:
    with pytest.raises(MotorError, match="inicio_contagem"):
        _multa("10.00", date(2025, 10, 31), date(2025, 11, 5), multa_mora_inicio_contagem="x")
    with pytest.raises(MotorError, match="arredondamento"):
        _multa("10.00", date(2025, 10, 31), date(2025, 11, 5),
               multa_mora_arredondamento="ROUND_HALF_EVEN")  # fmt: skip
