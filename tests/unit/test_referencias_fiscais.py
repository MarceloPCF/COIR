"""Tabelas de data/ do motor fiscal: parâmetros com vigência, tipos de ativo e trocas."""

from datetime import date
from decimal import Decimal

import pytest

from coir import referencias
from coir.domain.fiscal import TipoAtivo
from coir.erros import ReferenciaError

FISCAIS = referencias.carregar_fiscais()
REFS = referencias.carregar()


def test_parametros_por_vigencia() -> None:
    p = FISCAIS.parametros
    assert p.decimal("limite_isencao_acoes", date(2020, 1, 31)) == Decimal("20000.00")
    assert p.booleano("day_trade_mesma_corretora", date(2010, 7, 27)) is False
    assert p.booleano("day_trade_mesma_corretora", date(2010, 7, 28)) is True
    with pytest.raises(ReferenciaError, match="sem vigência"):
        p.decimal("aliquota_comum", date(2004, 12, 31))


def test_tipos_de_ativo() -> None:
    t = FISCAIS.tipos
    assert (t.tipo("PETR4"), t.tipo("B3SA3"), t.tipo("AAPL34")) == (
        TipoAtivo.ACAO, TipoAtivo.ACAO, TipoAtivo.BDR)  # fmt: skip
    assert (t.tipo("TAEE11"), t.tipo("BOVA11"), t.tipo("B5P211")) == (
        TipoAtivo.UNIT, TipoAtivo.ETF, TipoAtivo.ETF_RF)  # fmt: skip
    assert t.tipo("XPML11") is TipoAtivo.FII  # pela tabela gerada do cadastro da B3 (F7b)
    with pytest.raises(ReferenciaError, match=r"tipos_ativo.csv"):
        t.tipo("XYZW11")  # sem cadastro: sem palpite


def test_serie_de_opcao_segue_a_troca_do_ativo_objeto() -> None:
    """3R Petroleum -> Brava (RRRP3 -> BRAV3 a partir de 09/09/2024): a série vendida como
    RRRPT32 foi exercida como BRAVT32."""
    assert REFS.ticker_na_data("RRRPT32", date(2023, 12, 18)) == "BRAVT32"
    assert REFS.ticker_na_data("RRRP3", date(2024, 9, 6)) == "BRAV3"
    assert REFS.ticker_na_data("RRRP3", date(2024, 9, 9)) == "RRRP3"
    assert REFS.ticker_na_data("PETRA10", date(2023, 1, 2)) == "PETRA10"


def test_isencao_pme_por_vigencia_da_compra() -> None:
    prio = FISCAIS.pme["PRIO3"]
    assert prio.compra_elegivel(date(2020, 6, 1)) and not prio.compra_elegivel(date(2021, 2, 1))
    assert prio.vendas_ate == date(2023, 12, 31)


def test_calendario_cobre_o_ano_seguinte() -> None:
    """Revisão S1: o vencimento da DARF de dezembro e a multa precisam dos feriados do ano
    seguinte; sem eles, as telas de imposto falham. Este teste avisa com um ano de folga."""
    from datetime import date

    from coir import referencias

    assert referencias.carregar_fiscais().calendario.ultimo_ano >= date.today().year + 1
