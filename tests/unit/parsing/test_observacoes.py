"""Coluna Obs. (*) da linha de negócio: marcas independentes, em qualquer ordem, em qualquer layout.

Legenda do rodapé das notas SINACOR: 2 corretora ou pessoa vinculada na contraparte, # negócio
direto, 8 liquidação institucional, D day trade, F cobertura, B debêntures, A posição futuro, T
liquidação pelo bruto, C clubes e fundos, I POP, P carteira própria, H home broker, X box, Y
desmanche de box, L precatório. "@" e "4" aparecem nas notas, sem constar da legenda.
"""

from itertools import permutations
from pathlib import Path

import pytest
from hypothesis import given
from hypothesis import strategies as st

from coir import referencias
from coir.domain.nota import NegocioBovespa
from coir.erros import ReferenciaError
from coir.parsing import bovespa, observacoes
from coir.parsing.extrair import extrair_textos
from coir.parsing.normalizar import operacoes

PASTA = Path(__file__).parents[2] / "fixtures" / "notas"
REFS = referencias.carregar()


def _linha(espec: str, obs: str, cv: str = "C") -> NegocioBovespa:
    dc = "D" if cv == "C" else "C"
    [n] = bovespa.negocios(f"1-BOVESPA {cv} VISTA {espec} {obs} 100 10,00 1.000,00 {dc}", "xp")
    return n


@pytest.mark.parametrize("obs", ["D", "D#", "#D", "D@", "@D", "D#@", "@#D", "#@D", "D#2", "2#D",
                                 "HD", "DH", "H#D", "D8", "8D", "@D#2", "HD#8"])  # fmt: skip
@pytest.mark.parametrize("espec", ["PETZCOBASI ON NM", "RAIZEN PNA P N2", "XPETFCHINA CI",
                                   "INVESTO LFTB F11", "ALUPAR UNT N2", "PETR4",
                                   "PETR4F"])  # fmt: skip
def test_marca_d_em_qualquer_ordem_e_em_qualquer_ativo(espec: str, obs: str) -> None:
    n = _linha(espec, obs)
    assert (n.especificacao, n.observacao, n.day_trade) == (espec, obs, True)


@given(st.permutations("#@D28H").flatmap(lambda p: st.integers(1, 6).map(lambda k: "".join(p[:k]))))
def test_qualquer_combinacao_com_simbolo_ou_em_classe_nunca_fica_no_nome(obs: str) -> None:
    n = _linha("VALE ON NM", obs)
    assert n.especificacao == "VALE ON NM"
    assert n.observacao == obs
    assert n.day_trade is ("D" in obs)


@pytest.mark.parametrize(
    "marca", sorted(observacoes.MARCAS_DA_LEGENDA | observacoes.MARCAS_SEM_LEGENDA)
)
def test_cada_marca_da_legenda_sozinha(marca: str) -> None:
    n = _linha("VALE ON NM", marca)
    assert (n.especificacao, n.observacao, n.day_trade) == ("VALE ON NM", marca, marca == "D")


@pytest.mark.parametrize("par", [("D", "#"), ("D", "@"), ("H", "#"), ("#", "2")])
def test_as_duas_ordens_dao_o_mesmo_ativo_e_a_mesma_marca(par: tuple[str, str]) -> None:
    a, b = (_linha("VALE ON NM", "".join(p)) for p in permutations(par))
    assert (a.especificacao, a.day_trade) == (b.especificacao, b.day_trade)


@pytest.mark.parametrize(
    ("espec", "obs"),
    [("XPETFCHINA CI", ""), ("BB ETF SP DV CI", ""), ("ALLIANCE DAT", ""), ("BANCO X DB", ""),
     ("FUNDO CIA", ""), ("VENTURA IBX", "")],
)  # fmt: skip
def test_especificacao_com_letras_da_legenda_nao_vira_marca(espec: str, obs: str) -> None:
    n = _linha(espec, obs)
    assert (n.especificacao, n.observacao, n.day_trade) == (espec, "", False)


def test_letra_ambigua_e_marca_so_depois_de_classe_ou_sufixo() -> None:
    assert _linha("VALE ON NM", "C").especificacao == "VALE ON NM"
    assert _linha("NOME FUNDO", "C").especificacao == "NOME FUNDO C"
    assert _linha("RAIZEN PNA P", "N2").especificacao == "RAIZEN PNA P N2"


def test_marca_depois_de_especificacao_colada() -> None:
    """Layout antigo da XP: "ON NM" sai colado na linha."""
    n = _linha("NOME ONNM", "HD")
    assert (n.especificacao, n.observacao, n.day_trade) == ("NOME ONNM", "HD", True)


def test_lote_de_mil_junto_de_marca() -> None:
    [n] = bovespa.negocios("1-BOVESPA C VISTA GERDAU PN * D# 1.000 10,00 10,00 D", "xp")
    assert (n.especificacao, n.observacao, n.lote_mil) == ("GERDAU PN", "D#", True)
    assert n.day_trade


def test_lote_de_mil_depois_das_marcas_nao_as_prende() -> None:
    [n] = bovespa.negocios("1-BOVESPA C VISTA GERDAU PN D * 1.000 10,00 10,00 D", "xp")
    assert (n.especificacao, n.observacao, n.lote_mil) == ("GERDAU PN", "D", True)


def test_a_legenda_da_propria_nota_vale() -> None:
    """Letra nova explicada no rodapé da nota é marca daquela nota."""
    texto = "(*) Observações A - Posição futuro\nQ - Nova marca\n"
    assert {"A", "Q"} <= observacoes.legenda(texto)
    linha = "1-BOVESPA C VISTA VALE ON NM Q 100 10,00 1.000,00 D\n"
    assert bovespa.negocios(linha + texto, "xp")[0].observacao == "Q"
    assert bovespa.negocios(linha, "xp")[0].especificacao == "VALE ON NM Q"


@pytest.mark.parametrize("fixture", sorted(p.stem for p in PASTA.glob("*bovespa*.txt")))
def test_legenda_lida_em_todos_os_layouts_tem_o_day_trade(fixture: str) -> None:
    texto = (PASTA / f"{fixture}.txt").read_text(encoding="utf-8")
    if observacoes.RE_INICIO_LEGENDA.search(texto):
        assert "D" in observacoes.legenda(texto)


def test_layout_com_codigo_na_linha_nao_perde_o_day_trade() -> None:
    """Falha silenciosa antiga: "PETR4 D@" ficava com o código certo e sem day trade."""
    n = _linha("PETR4", "D@")
    assert n.day_trade and n.observacao == "D@" and n.especificacao == "PETR4"


def test_marca_desconhecida_com_d_recusa_em_vez_de_virar_operacao_comum() -> None:
    nota = extrair_textos("x.pdf", (PASTA / "xp_2026_08_bovespa.txt").read_text(
        encoding="utf-8").split("\f")).notas[0]  # fmt: skip
    for espec in ("PETR4 DZ", "VALE ON NM DZ"):
        estranho = _linha(espec, "")
        nova = nota.model_copy(update={"negocios_bovespa": (estranho,)})
        with pytest.raises(ReferenciaError, match="marca de day trade"):
            operacoes(nova, REFS, {})


def test_codigo_na_linha_com_d_no_meio_de_marca_nova_nao_passa_em_silencio() -> None:
    nota = extrair_textos("x.pdf", (PASTA / "xp_2026_08_bovespa.txt").read_text(
        encoding="utf-8").split("\f")).notas[0]  # fmt: skip
    nova = nota.model_copy(update={"negocios_bovespa": (_linha("PETR4 D Z", ""),)})
    with pytest.raises(ReferenciaError, match="marca de day trade"):
        operacoes(nova, REFS, {})


def test_fim_do_nome_junto_de_outra_marca() -> None:
    """Nome real da B3 terminado em letra da legenda, com o D em seguida: "GOOGLE DRN A" + D."""
    nota = extrair_textos("x.pdf", (PASTA / "xp_2026_08_bovespa.txt").read_text(
        encoding="utf-8").split("\f")).notas[0]  # fmt: skip
    refs = REFS.model_copy(update={"nomes_historicos": {
        "googledrna": (("2010-01-01", "2030-01-01", "GOGL35"),)}})  # fmt: skip
    negocio = _linha("GOOGLE DRN", "AD")
    assert negocio.observacao == "AD" and negocio.day_trade
    [op] = operacoes(nota.model_copy(update={"negocios_bovespa": (negocio,)}), refs, {})
    assert (op.papel, op.categoria) == ("GOGL35", "DAY_TRADE")


def test_marca_nova_sem_d_nao_trava_a_nota() -> None:
    nota = extrair_textos("x.pdf", (PASTA / "xp_2026_08_bovespa.txt").read_text(
        encoding="utf-8").split("\f")).notas[0]  # fmt: skip
    nova = nota.model_copy(update={"negocios_bovespa": (_linha("VALE ON NM ZK", ""),)})
    [op] = operacoes(nova, REFS, {})
    assert (op.papel, op.categoria) == ("VALE3", "COMUM")


def test_box_entra_com_aviso() -> None:
    res = extrair_textos("x.pdf", (PASTA / "agora_2008_04_bovespa.txt").read_text(
        encoding="utf-8").replace("N1 H 300", "N1 X 300", 1).split("\f"))  # fmt: skip
    assert len(res.notas) == 1
    assert any("marca X (box)" in a for a in res.avisos)
