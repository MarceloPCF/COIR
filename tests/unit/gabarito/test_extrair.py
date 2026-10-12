from datetime import date
from decimal import Decimal

import pytest

from coir.erros import GabaritoError
from coir.gabarito.extrair import extrair
from coir.gabarito.modelo import Categoria, Gabarito
from coir.gabarito.pseudonimos import MapaPseudonimos
from tests.unit.gabarito import fabrica
from tests.unit.gabarito.fabrica import CPF_A, CPF_A_ERRADO


def _extrair() -> tuple[Gabarito, MapaPseudonimos]:
    mapa = MapaPseudonimos()
    return extrair(fabrica.planilha(), mapa), mapa


def test_operacoes_com_nota_e_eventos_separados() -> None:
    g, _ = _extrair()
    assert [(o.categoria, o.nota) for o in g.operacoes] == [
        (Categoria.COMUM, "2588033"),
        (Categoria.COMUM, "11490605"),
        (Categoria.DAY_TRADE, "400729"),
    ]
    assert [(e.tipo, e.papel, e.quantidade) for e in g.eventos] == [
        ("Desdobramento", "GGBR4", Decimal("300.00000000"))
    ]
    assert len(g.eventos_cadastro) == 1


def test_valores_normalizados_e_pseudonimizados() -> None:
    g, mapa = _extrair()
    primeira = g.operacoes[0]
    assert primeira.corretora == "AGORA"
    assert primeira.conta == "AGORA-1"
    assert primeira.titular == "TITULAR-1"
    assert primeira.data == date(2008, 4, 18)
    assert primeira.custos == Decimal("26.92000000")
    assert g.operacoes[1].corretora == "XP"  # "XP " normalizado
    assert mapa.titulares == {CPF_A: "TITULAR-1"}


def test_cpf_digitado_errado_vira_titular_principal_com_aviso_sem_cpf() -> None:
    g, _ = _extrair()
    assert g.eventos[0].titular == "TITULAR-1"
    avisos = " ".join(g.avisos)
    assert "erro de digitação" in avisos
    assert CPF_A_ERRADO not in avisos and CPF_A not in avisos


def test_posicoes_ignoram_linhas_sem_tipo() -> None:
    g, _ = _extrair()
    comuns = [p for p in g.posicoes if p.categoria is Categoria.COMUM]
    assert [
        (p.ordem, p.cv, p.quantidade_apos, p.resultado, p.mes_competencia, p.isento) for p in comuns
    ] == [
        (1, "C", Decimal("300"), Decimal("0"), 4, False),
        (2, "V", Decimal("200"), Decimal("2015.74"), 6, True),
    ]
    assert [p.mercado for p in g.posicoes if p.categoria is Categoria.DAY_TRADE] == ["INDICE"]


def test_apuracao_por_ano_ate_o_ultimo_ano_com_operacao() -> None:
    g, _ = _extrair()
    comum = [a for a in g.apuracao if a.categoria is Categoria.COMUM]
    assert len(comum) == 12  # 2009 fica de fora: não há operação depois de 2008
    maio = next(a for a in comum if a.mes == 5)
    assert (maio.resultado, maio.irrf, maio.a_pagar, maio.base_calculo) == (
        Decimal("2015.74"),
        Decimal("0.41"),
        Decimal("-0.41"),
        Decimal("8268"),
    )


def test_darf_isentos_e_avisos() -> None:
    g, _ = _extrair()
    assert [(d.ano, d.mes, d.codigo, d.valor, d.vencimento) for d in g.darfs] == [
        (2008, 6, "6015", Decimal("302.36000000"), date(2008, 7, 31))
    ]
    assert [(i.mes, i.valor) for i in g.isentos if i.valor] == [(5, Decimal("100.00000000"))]
    assert {i.ano for i in g.isentos} == {2008}
    assert any("BTC_Outros" in a for a in g.avisos)
    assert g.custodia_inicial == []


def test_coluna_ausente_falha_alto() -> None:
    abas = fabrica.planilha()
    abas["Normais_Dados"][0] = [c if c != "Custos" else "Custo" for c in abas["Normais_Dados"][0]]
    with pytest.raises(GabaritoError, match="Custos"):
        extrair(abas, MapaPseudonimos())


def test_cv_invalido_falha_alto() -> None:
    abas = fabrica.planilha()
    abas["Normais_Dados"][1][5] = "X"
    with pytest.raises(GabaritoError, match="C/V"):
        extrair(abas, MapaPseudonimos())
