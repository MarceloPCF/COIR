from datetime import date
from decimal import Decimal
from pathlib import Path

from coir import referencias
from coir.domain.parametros import IsencaoPME
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
    assert all(o.irrf == 0 for o in cand.operacoes if o.cv == "C")


def test_irrf_so_nas_vendas_tributaveis() -> None:
    """0,005% nas vendas comuns, sem day trade nem PME (Lei 13.043/2014 art. 16)."""
    paginas = (PASTA / "xp_2026_08_bovespa.txt").read_text(encoding="utf-8").split("\f")
    [nota] = extrair_textos("x.pdf", paginas).notas
    ops = operacoes(nota, referencias.carregar(), {})
    vendas = [o for o in ops if o.cv == "V" and o.categoria == "COMUM"]
    assert len({o.papel for o in vendas}) > 1
    papel = vendas[0].papel
    limite = date(2030, 1, 1)
    pme = {papel: IsencaoPME(ticker=papel, aquisicao_desde=date(2014, 7, 10),
                             aquisicao_ate=limite, vendas_ate=limite, fonte="x")}  # fmt: skip
    cand = montar([(nota, ops)], MapaPseudonimos(), pme=pme)
    assert sum(o.irrf for o in cand.operacoes) == Decimal("4.20")
    assert all(o.irrf == 0 for o in cand.operacoes if o.papel == papel)
    assert all(o.irrf == 0 for o in cand.operacoes if o.categoria == "DAY_TRADE" or o.cv == "C")
    assert sum(o.total for o in cand.operacoes) == Decimal("84068.00")
