from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
from hypothesis import given
from hypothesis import strategies as st

from coir.erros import GabaritoError
from coir.gabarito.comparar import Tolerancias, comparar, comparar_apuracao, comparar_operacoes
from coir.gabarito.extrair import extrair
from coir.gabarito.modelo import Gabarito, OperacaoGabarito
from coir.gabarito.pseudonimos import MapaPseudonimos
from tests.unit.gabarito import fabrica

TOL = Tolerancias(operacao=Decimal("0.01"), pm=Decimal("0.05"), mes=Decimal("1.00"))


def _gabarito() -> Gabarito:
    return extrair(fabrica.planilha(), MapaPseudonimos())


def _troca_op(g: Gabarito, i: int, **campos: object) -> Gabarito:
    ops = list(g.operacoes)
    ops[i] = ops[i].model_copy(update=campos)
    return g.model_copy(update={"operacoes": ops})


def test_tolerancias_do_arquivo_do_projeto() -> None:
    assert Tolerancias.carregar() == TOL


def test_tolerancias_invalidas(tmp_path: Path) -> None:
    arquivo = tmp_path / "t.toml"
    arquivo.write_text('operacao = "0.01"\n', encoding="utf-8")
    with pytest.raises(GabaritoError):
        Tolerancias.carregar(arquivo)


def test_identico_nao_diverge() -> None:
    g = _gabarito()
    rel = comparar(g, g, TOL)
    assert rel.ok
    assert rel.comparados == {"operacoes": 3, "posicoes": 3, "apuracao": 24}


def test_operacao_dentro_e_fora_da_tolerancia() -> None:
    g = _gabarito()
    total = g.operacoes[0].total
    assert comparar(g, _troca_op(g, 0, total=total + Decimal("0.01")), TOL, ["operacoes"]).ok
    rel = comparar(g, _troca_op(g, 0, total=total + Decimal("0.02")), TOL, ["operacoes"])
    assert [(d.campo, d.diferenca) for d in rel.divergencias] == [("total", Decimal("0.02000000"))]


def test_quantidade_tem_que_ser_exata() -> None:
    g = _gabarito()
    rel = comparar(g, _troca_op(g, 0, quantidade=Decimal("299")), TOL, ["operacoes"])
    assert [d.campo for d in rel.divergencias] == ["quantidade"]


def test_consolidacao_diferente_nao_gera_divergencia() -> None:
    """O gabarito às vezes junta duas linhas da mesma nota; o candidato pode separá-las."""
    g = _gabarito()
    op = g.operacoes[0]
    metade = {k: getattr(op, k) / 2 for k in ("quantidade", "total", "custos", "irrf")}
    partes: list[OperacaoGabarito] = [op.model_copy(update=metade), op.model_copy(update=metade)]
    candidato = g.model_copy(update={"operacoes": partes + list(g.operacoes[1:])})
    assert comparar(g, candidato, TOL, ["operacoes"]).ok


def test_faltando_e_sobrando() -> None:
    g = _gabarito()
    menos = g.model_copy(update={"operacoes": g.operacoes[1:]})
    outra = g.operacoes[0].model_copy(update={"nota": "999", "data": date(2008, 4, 19)})
    mais = g.model_copy(update={"operacoes": [*g.operacoes, outra]})
    assert len(comparar(g, menos, TOL, ["operacoes"]).faltando) == 1
    assert len(comparar(g, mais, TOL, ["operacoes"]).sobrando) == 1


def test_pm_e_mes_nos_limites() -> None:
    g = _gabarito()
    pos = list(g.posicoes)
    pos[1] = pos[1].model_copy(update={"pm_apos": pos[1].pm_apos + Decimal("0.05")})  # type: ignore[operator]
    assert comparar(g, g.model_copy(update={"posicoes": pos}), TOL, ["posicoes"]).ok
    pos[1] = pos[1].model_copy(update={"pm_apos": pos[1].pm_apos + Decimal("0.01")})  # type: ignore[operator]
    rel = comparar(g, g.model_copy(update={"posicoes": pos}), TOL, ["posicoes"])
    assert [d.campo for d in rel.divergencias] == ["pm_apos"]

    meses = list(g.apuracao)
    meses[4] = meses[4].model_copy(update={"resultado": meses[4].resultado + Decimal("1.01")})
    rel = comparar(g, g.model_copy(update={"apuracao": meses}), TOL, ["apuracao"])
    assert [(d.chave, d.campo) for d in rel.divergencias] == [("('COMUM', 2008, 5)", "resultado")]


def test_nivel_desconhecido() -> None:
    g = _gabarito()
    with pytest.raises(GabaritoError):
        comparar(g, g, TOL, ["outro"])


centavos = st.decimals(min_value=Decimal("-0.01"), max_value=Decimal("0.01"), places=2)


@given(delta=centavos)
def test_propriedade_variacao_dentro_da_tolerancia_nunca_acusa(delta: Decimal) -> None:
    g = _gabarito()
    candidato = _troca_op(
        g, 1, total=g.operacoes[1].total + delta, custos=g.operacoes[1].custos - delta
    )
    rel = comparar(g, candidato, TOL, ["operacoes"])
    assert rel.ok


@given(valor=st.decimals(min_value=Decimal("-1e9"), max_value=Decimal("1e9"), places=8))
def test_propriedade_comparar_consigo_mesmo(valor: Decimal) -> None:
    g = _gabarito()
    g2 = _troca_op(g, 0, total=valor)
    from coir.gabarito.comparar import Relatorio

    rel = Relatorio()
    comparar_operacoes(g2.operacoes, g2.operacoes, TOL, rel)
    comparar_apuracao(g2.apuracao, g2.apuracao, TOL, rel)
    assert rel.ok


def test_pm_vazio_so_de_um_lado_diverge() -> None:
    g = _gabarito()
    pos = list(g.posicoes)
    assert pos[1].pm_apos is not None
    pos[1] = pos[1].model_copy(update={"pm_apos": None})
    rel = comparar(g, g.model_copy(update={"posicoes": pos}), TOL, ["posicoes"])
    assert [(d.campo, d.obtido) for d in rel.divergencias] == [("pm_apos", None)]
