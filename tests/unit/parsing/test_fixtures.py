"""As 10 notas anonimizadas de tests/fixtures/notas extraídas e conciliadas = o esperado.json."""

import json
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from coir.domain.nota import Segmento
from coir.parsing.extrair import extrair_textos

PASTA = Path(__file__).parents[2] / "fixtures" / "notas"
FIXTURES = sorted(p.name.removesuffix(".txt") for p in PASTA.glob("*.txt"))


def _ler(nome: str) -> tuple[list[str], dict]:
    texto = (PASTA / f"{nome}.txt").read_text(encoding="utf-8")
    esperado = json.loads((PASTA / f"{nome}.esperado.json").read_text(encoding="utf-8"))
    return texto.split("\f"), esperado


def test_ha_dez_fixtures() -> None:
    assert len(FIXTURES) == 10


@pytest.mark.parametrize("nome", FIXTURES)
def test_fixture_extraida_conciliada_e_igual_ao_esperado(nome: str) -> None:
    paginas, esperado = _ler(nome)
    res = extrair_textos(f"{nome}.pdf", paginas)
    assert res.rejeitadas == []
    [nota] = res.notas
    assert nota.corretora == esperado["corretora"].upper()
    assert nota.segmento is Segmento(esperado["segmento"])
    assert nota.numero == esperado["numero_nota"]
    assert nota.data_pregao == date.fromisoformat(esperado["data_pregao"])
    assert len(nota.paginas) == esperado.get("folhas", 1)
    ops = esperado["operacoes"]
    lidos = nota.negocios_bovespa or nota.negocios_bmf
    assert [(n.cv, n.quantidade, n.preco, n.valor) for n in lidos] == [
        (o["cv"], Decimal(o["quantidade"]), Decimal(o["preco"]), Decimal(o["valor"])) for o in ops
    ]
    liquido = esperado["resumo_financeiro"]["liquido"]
    resumo = nota.resumo_bovespa or nota.resumo_bmf
    assert resumo is not None
    assert (resumo.liquido.valor, resumo.liquido.dc) == (Decimal(liquido["valor"]), liquido["dc"])


def test_bovespa_marcas_e_especificacao() -> None:
    paginas, _ = _ler("xp_2016_11_bovespa_antigo_multipagina")
    [nota] = extrair_textos("x.pdf", paginas).notas
    assert {n.observacao for n in nota.negocios_bovespa} >= {"D", "#"}
    assert any(n.day_trade for n in nota.negocios_bovespa)
    r = nota.resumo_bovespa
    assert r is not None and r.irrf == Decimal("2.23") and r.irrf_no_liquido


def test_btg_opcoes_irrf_nao_descontado_e_exercicio() -> None:
    paginas, _ = _ler("btg_2023_01_bovespa_opcoes")
    [nota] = extrair_textos("x.pdf", paginas).notas
    r = nota.resumo_bovespa
    assert r is not None and (r.irrf, r.irrf_no_liquido) == (Decimal("0.01"), False)
    assert nota.negocios_bovespa[0].prazo == "02/23"
    paginas, _ = _ler("btg_2023_01_bovespa_exercicio")
    [ex] = extrair_textos("x.pdf", paginas).notas
    assert ex.negocios_bovespa[0].mercado == "EXERC OPC VENDA"
    assert ex.negocios_bovespa[0].especificacao == "KLBNM199E"


def test_bmf_custos_ajustes_e_irrf_day_trade() -> None:
    paginas, _ = _ler("xp_2019_02_bmf_daytrade")
    [nota] = extrair_textos("x.pdf", paginas).notas
    r = nota.resumo_bmf
    assert r is not None
    assert (r.ajuste_day_trade.valor, r.ajuste_day_trade.dc) == (Decimal("44.00"), "C")
    assert (r.total_custos, r.impostos, r.irrf_day_trade) == (
        Decimal("2.64"),
        Decimal("0.08"),
        Decimal("0.41"),
    )
    assert nota.custos_total == Decimal("2.72")
    paginas, _ = _ler("xp_2024_01_bmf_novo_ajupos_negativo")
    [ajupos] = extrair_textos("x.pdf", paginas).notas
    assert any(n.quantidade < 0 and n.tipo == "AJUPOS" for n in ajupos.negocios_bmf)
