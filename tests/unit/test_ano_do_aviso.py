"""Ano de cada aviso da apuração: a tela Apuração mostra só os do ano escolhido, e o sino diz em
que anos há aviso. Os avisos começam com a data em três formatos (ou sem data)."""

import pytest

from coir.domain.fiscal import Apuracao


@pytest.mark.parametrize(
    ("aviso", "ano"),
    [
        ("2026-03-10: venda de AAAA3 maior que a posição", 2026),
        ("03/2026: as vendas de ações em day trade (R$ 8.000,00) entram na soma", 2026),
        ("10/03/2025: bonificação de AAAA3 com valor atribuído", 2025),
        ("2019: saldo inicial", 2019),
        ("AAAA3: sem tipo de ativo nas tabelas", None),
        ("aviso sem data", None),
        ("", None),
    ],
)
def test_ano_do_aviso(aviso: str, ano: int | None) -> None:
    """Antes, só "AAAA-..." era reconhecido: os avisos "MM/AAAA" apareciam em todos os anos na tela
    Apuração e o sino não dizia o ano deles (revisão fiscal de 12/11/2026)."""
    assert Apuracao.ano_do_aviso(aviso) == ano
