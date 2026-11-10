"""Custos da nota que vão para o motor: a taxa de custódia fica fora (SC Cosit 227/2018 §§17-18;
parâmetro custodia_no_custo, decisão do usuário de 10/11/2026, leitura conservadora)."""

from decimal import Decimal
from pathlib import Path

from coir.domain.nota import CategoriaCusto
from coir.parsing.extrair import extrair_textos
from coir.servicos.apuracao import custos_dedutiveis
from tests.fiscal_cases.apoio import parametros

NOTA = Path(__file__).parents[1] / "fixtures" / "notas" / "xp_2026_08_bovespa.txt"


def test_custodia_sai_do_custo_e_volta_com_o_parametro() -> None:
    [nota] = extrair_textos("x.pdf", NOTA.read_text(encoding="utf-8").split("\f")).notas
    resumo = nota.resumo_bovespa
    assert resumo is not None
    # a nota fictícia traz custódia zero: põe R$ 5,00 nela, sem mexer no resto
    custos = [
        c.model_copy(update={"valor": Decimal("5.00")})
        if c.categoria is CategoriaCusto.CUSTODIA
        else c
        for c in resumo.custos
    ]
    assert any(c.categoria is CategoriaCusto.CUSTODIA for c in custos)
    nota = nota.model_copy(update={"resumo_bovespa": resumo.model_copy(update={"custos": custos})})
    total = nota.custos_total
    assert custos_dedutiveis(nota, total, parametros()) == total - Decimal("5.00")
    assert custos_dedutiveis(nota, total, parametros(custodia_no_custo="true")) == total
