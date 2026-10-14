from decimal import Decimal
from pathlib import Path

from coir import referencias
from coir.gabarito.candidato import _distribuir, montar
from coir.gabarito.pseudonimos import MapaPseudonimos
from coir.parsing.extrair import extrair_textos
from coir.parsing.normalizar import operacoes

PASTA = Path(__file__).parents[2] / "fixtures" / "notas"


def test_distribuir_soma_exata() -> None:
    partes = _distribuir(Decimal("20.40"), [Decimal(1), Decimal(1), Decimal(1)])
    assert sum(partes) == Decimal("20.40")
    assert _distribuir(Decimal("1"), [Decimal(0), Decimal(0)]) == [Decimal(0), Decimal(1)]
    assert _distribuir(Decimal("1"), []) == []


def test_candidato_pseudonimizado_com_totais_da_nota() -> None:
    paginas = (PASTA / "xp_2026_08_bovespa.txt").read_text(encoding="utf-8").split("\f")
    [nota] = extrair_textos("x.pdf", paginas).notas
    ops = operacoes(nota, referencias.carregar(), {})
    mapa = MapaPseudonimos()
    cand = montar([(nota, ops)], mapa)
    assert {o.conta for o in cand.operacoes} == {"XP-1"} and mapa.contas == {"XP|0": "XP-1"}
    assert sum(o.custos for o in cand.operacoes) == nota.custos_total == Decimal("36.02")
    assert sum(o.irrf for o in cand.operacoes) == Decimal("4.20")
    assert sum(o.total for o in cand.operacoes) == Decimal("84068.00")
