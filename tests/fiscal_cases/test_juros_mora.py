"""Juros de mora da DARF 6015 paga em atraso (números conferidos à mão em 19/10/2026).

Fontes: Lei 9.430/1996 art. 61 §3º c/c art. 5º §3º (Selic acumulada mensalmente + 1% no mês do
pagamento) e art. 68 §1º; RFB, "Informações úteis sobre o Sicalc" (soma da Selic desde a do mês
seguinte ao do vencimento até a do mês anterior ao do pagamento, mais 1%; pago no próprio mês do
vencimento, sem juros; base = o tributo devido); glossário do Sicalc (DAC: data não útil vale no
próximo dia útil); BCB SGS 4390 (Selic acumulada no mês, a mesma da consulta Selic do Sicalc).
Centavos truncados: conferido no Sicalc pelo usuário (17/10/2026: 12,3457 -> 12,34).
"""

from datetime import date
from decimal import Decimal

import pytest

from coir.engine.darf import juros_mora, multa_mora
from coir.erros import MotorError
from tests.fiscal_cases.apoio import REFS

D = Decimal
PARAMS, CAL, SELIC = REFS.parametros, REFS.calendario, REFS.selic


def _juros(
    principal: str,
    vencimento: date,
    pagamento: date,
    selic: dict[tuple[int, int], Decimal] | None = None,
) -> tuple[Decimal, Decimal]:
    j = juros_mora(
        D(principal),
        date(vencimento.year, 1, 1),
        vencimento,
        pagamento,
        PARAMS,
        CAL,
        SELIC if selic is None else selic,
    )
    return j.percentual, j.valor


def _multa(principal: str, vencimento: date, pagamento: date) -> Decimal:
    return multa_mora(
        D(principal), date(vencimento.year, 1, 1), vencimento, pagamento, PARAMS, CAL
    ).valor


def test_j1_caso_conferido_no_sicalc() -> None:
    """6015, comp. 08/2026, R$ 1.234,57, vence qua 30/09/2026, DAC qua 07/10/2026: nenhum mês de
    Selic entre os dois, só 1% = 12,3457 -> 12,34; multa 28,51; total 1.275,42 (Sicalc)."""
    percentual, juros = _juros("1234.57", date(2026, 9, 30), date(2026, 10, 7))
    multa = _multa("1234.57", date(2026, 9, 30), date(2026, 10, 7))
    assert (percentual, juros, multa) == (D("0.01"), D("12.34"), D("28.51"))
    assert D("1234.57") + multa + juros == D("1275.42")


def test_j2_varios_meses_atravessando_a_virada_do_ano() -> None:
    """Comp. 10/2025, R$ 3.456,78, vence sex 28/11/2025, DAC seg 16/03/2026: Selic 12/2025 1,22 +
    01/2026 1,16 + 02/2026 1,00 = 3,38 + 1% = 4,38% (tabela RFB de mar/2026 para vencimento
    11/2025) -> 151,406964 -> 151,40; multa no teto (106 dias) 691,35; total 4.299,53."""
    percentual, juros = _juros("3456.78", date(2025, 11, 28), date(2026, 3, 16))
    multa = _multa("3456.78", date(2025, 11, 28), date(2026, 3, 16))
    assert (percentual, juros, multa) == (D("0.0438"), D("151.40"), D("691.35"))
    assert D("3456.78") + multa + juros == D("4299.53")


def test_j3_data_nao_util_muda_o_mes() -> None:
    """R$ 800,00, vence qui 31/07/2025, pagamento informado dom 31/08/2025 -> DAC seg 01/09/2025:
    entra a Selic de 08/2025 (1,16) + 1% = 2,16% = 17,28; multa 32 dias 10,56% = 84,48; total
    901,76 (sem o ajuste da data seriam 8,00 de juros e 81,84 de multa)."""
    j = juros_mora(
        D("800.00"), date(2025, 1, 1), date(2025, 7, 31), date(2025, 8, 31), PARAMS, CAL, SELIC
    )
    assert (j.pagamento, j.percentual, j.valor) == (date(2025, 9, 1), D("0.0216"), D("17.28"))
    assert j.taxas == (("08/2025", D("1.16")),)
    assert _multa("800.00", date(2025, 7, 31), date(2025, 8, 31)) == D("84.48")


def test_j4_atraso_longo_com_duas_viradas_de_ano() -> None:
    """Comp. 03/2024, R$ 4.321,59, vence ter 30/04/2024, DAC qui 15/10/2026: Selic de 05/2024 a
    09/2026 = 6,89 + 13,47 + 10,04 = 30,40 + 1% = 31,40% -> 1.356,97926 -> 1.356,97; multa no teto
    864,31; total 6.542,87."""
    percentual, juros = _juros("4321.59", date(2024, 4, 30), date(2026, 10, 15))
    assert (percentual, juros) == (D("0.3140"), D("1356.97"))
    assert D("4321.59") + _multa("4321.59", date(2024, 4, 30), date(2026, 10, 15)) + juros == D(
        "6542.87"
    )


def test_j5_mes_sem_selic_publicada_e_estimativa() -> None:
    """Mesmo DARF do J1 com DAC qui 05/11/2026 e a tabela até 09/2026: a Selic de 10/2026 falta,
    os juros saem com o 1% (12,34) e o mês aparece como estimativa."""
    ate_setembro = {m: t for m, t in SELIC.items() if m <= (2026, 9)}
    j = juros_mora(
        D("1234.57"),
        date(2026, 1, 1),
        date(2026, 9, 30),
        date(2026, 11, 5),
        PARAMS,
        CAL,
        ate_setembro,
    )
    assert (j.valor, j.meses_sem_taxa) == (D("12.34"), ("10/2026",))
    assert _multa("1234.57", date(2026, 9, 30), date(2026, 11, 5)) == D("146.66")


def test_j6_pago_no_mes_do_vencimento_sem_juros() -> None:
    """Exemplo da RFB: vence 14/11/2025 e paga em 28/11/2025 (mesmo mês): juros zero, multa sim."""
    assert _juros("1000.00", date(2025, 11, 14), date(2025, 11, 28)) == (D(0), D(0))
    assert _multa("1000.00", date(2025, 11, 14), date(2025, 11, 28)) > 0
    assert _juros("1000.00", date(2025, 11, 14), date(2025, 11, 14)) == (D(0), D(0))


def test_j7_31_de_dezembro_sem_expediente_vai_para_janeiro() -> None:
    """R$ 700,00, vence ter 30/12/2025, pagamento informado qua 31/12/2025: no calendário do COIR
    31/12 não tem expediente bancário, a DAC vai a 02/01/2026: juros 1% = 7,00; multa 2,31."""
    assert _juros("700.00", date(2025, 12, 30), date(2025, 12, 31)) == (D("0.01"), D("7.00"))
    assert _multa("700.00", date(2025, 12, 30), date(2025, 12, 31)) == D("2.31")


@pytest.mark.parametrize(
    ("vencimento", "pagamento", "publicado"),
    [
        (date(2007, 12, 28), date(2021, 4, 15), "119.72"),
        (date(2008, 1, 31), date(2021, 4, 15), "118.79"),
        (date(2012, 1, 31), date(2026, 7, 15), "135.51"),
        (date(2024, 4, 30), date(2026, 3, 16), "23.52"),
        (date(2024, 4, 30), date(2026, 7, 15), "28.01"),
        (date(2025, 11, 28), date(2026, 3, 16), "4.38"),
        (date(2026, 6, 30), date(2026, 8, 14), "2.22"),
        (date(2026, 7, 31), date(2026, 8, 14), "1.00"),
        (date(2026, 8, 3), date(2026, 8, 14), "0.00"),
    ],
)
def test_tabela_selic_confere_com_a_tabela_de_acrescimos_da_rfb(
    vencimento: date, pagamento: date, publicado: str
) -> None:
    """Tabelas de acréscimos legais da RFB (abr/2021, mar/2026, jul/2026 e ago/2026): 1% + soma
    da Selic do CSV tem de dar o percentual publicado para cada vencimento."""
    assert _juros("100.00", vencimento, pagamento)[0] * 100 == D(publicado)


def test_buraco_na_tabela_falha_alto() -> None:
    sem_marco = {m: t for m, t in SELIC.items() if m != (2025, 3)}
    with pytest.raises(MotorError, match="03/2025 ausente"):
        _juros("100.00", date(2025, 1, 31), date(2025, 6, 2), sem_marco)
    with pytest.raises(MotorError, match="negativo"):
        _juros("-1", date(2025, 1, 31), date(2025, 6, 2))
