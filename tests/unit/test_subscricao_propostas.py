"""Propostas da movimentação da B3 para a subscrição: o crédito do direito, o exercício (crédito do
recibo) e a conversão do recibo em ação (crédito das ações), sobre o banco de demonstração.

A sequência de linhas é a que a B3 mostra: "Direito de Subscrição" (entrada), "Solicitação de
Subscrição" e "Direitos de Subscrição - Exercido" (saídas), "Recibo de Subscrição" (entrada e
saída no mesmo dia) e, semanas depois, "Atualização" de entrada na ação, com a quantidade do
recibo. Nenhuma delas traz o valor pago."""

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest

from coir import referencias
from coir.db.modelos import MovimentacaoB3
from coir.domain.fiscal import NaturezaEvento, TipoEvento, Visao
from coir.erros import ImportacaoError
from coir.servicos import eventos
from coir.servicos.apuracao import carteira_banco
from coir.servicos.b3 import importar_planilhas
from coir.servicos.eventos import (
    DadosEvento,
    _mapa_proposta,  # pyright: ignore[reportPrivateUsage]
)
from tests.unit.test_importador_b3 import CAB_MOV, planilha
from tools.banco_demo import criar

AGORA = datetime(2026, 10, 19, 12, 0, 0)
REFS, FISCAIS = referencias.carregar(), referencias.carregar_fiscais()
AGORA_B3 = "AGORA CORRETORA DE TITULOS E VALORES MOBILIARIOS S/A"
AGORA_CONTA = 1  # GGBR4: 300 ações (custo 18.687) de 2022 em diante
D = Decimal
DIREITO, RECIBO, ACAO = "GGBR2 - GERDAU S.A.", "GGBR10 - GERDAU S.A.", "GGBR4 - GERDAU S.A."


@pytest.fixture
def demo(tmp_path: Path) -> Path:
    banco = criar(tmp_path / "demo")
    arquivo = planilha(banco.parent / "movimentacao-2023.xlsx", {"Movimentação": [
        CAB_MOV,
        ["Credito", "03/04/2023", "Direito de Subscrição", DIREITO, AGORA_B3, 100, "-", "-"],
        ["Debito", "28/04/2023", "Solicitação de Subscrição", DIREITO, AGORA_B3, 85, "-", "-"],
        ["Debito", "28/04/2023", "Direitos de Subscrição - Exercido", DIREITO, AGORA_B3, 85, "-",
         "-"],
        ["Credito", "02/05/2023", "Recibo de Subscrição", RECIBO, AGORA_B3, 85, "-", "-"],
        ["Debito", "02/05/2023", "Recibo de Subscrição", RECIBO, AGORA_B3, 85, "-", "-"],
        ["Credito", "16/06/2023", "Atualização", ACAO, AGORA_B3, 85, "-", "-"],
        ["Credito", "20/06/2023", "Atualização", ACAO, AGORA_B3, 7, "-", "-"],  # outra coisa
    ]})  # fmt: skip
    importar_planilhas(banco, [arquivo], 1, AGORA)
    return banco


def _carteira(banco: Path, dia: date) -> dict[str, tuple[Decimal, Decimal]]:
    pos, _ = carteira_banco(banco, REFS, FISCAIS, dia, Visao.TITULAR)
    return {p.ticker: (p.quantidade, p.custo.quantize(D("0.01"))) for p in pos}


def _por_tipo(banco: Path) -> dict[TipoEvento, eventos.Proposta]:
    lista = eventos.propostas(banco, 1, REFS)
    assert len(lista) == 3  # a "Atualização" de 7 e as saídas não propõem nada
    return {p.dados.tipo: p for p in lista}


def test_as_tres_propostas_da_subscricao(demo: Path) -> None:
    p = _por_tipo(demo)
    recebido = p[TipoEvento.DIREITO_RECEBIDO].dados
    assert (recebido.data, recebido.ticker, recebido.cv, recebido.quantidade, recebido.valor) == (
        date(2023, 4, 3), "GGBR2", "C", 100, 0)  # fmt: skip
    exercicio = p[TipoEvento.EXERCICIO_DIREITO].dados
    assert (exercicio.data, exercicio.ticker, exercicio.quantidade) == (
        date(2023, 5, 2), "GGBR2", 85)  # fmt: skip
    assert (exercicio.ticker_destino, exercicio.quantidade_destino, exercicio.valor) == (
        "GGBR10", 85, 0)  # fmt: skip
    assert "valor pago" in p[TipoEvento.EXERCICIO_DIREITO].nota
    conversao = p[TipoEvento.INCORPORACAO].dados
    assert (conversao.data, conversao.ticker, conversao.quantidade) == (
        date(2023, 6, 16), "GGBR10", 85)  # fmt: skip
    assert (conversao.ticker_destino, conversao.quantidade_destino, conversao.natureza) == (
        "GGBR4", 85, NaturezaEvento.RECIBO_SUBSCRICAO)  # fmt: skip
    assert all(x.situacao == "pendente" and x.dados.conta == AGORA_CONTA for x in p.values())


def test_lancadas_as_tres_o_custo_do_direito_chega_a_acao(demo: Path) -> None:
    """85 ações subscritas a 20,00 = 1.700,00. Direitos recebidos sem custo: GGBR4 passa de 300
    a 385 ações, com 1.700,00 a mais de custo, só na conversão; até lá o recibo guarda o valor."""
    p = _por_tipo(demo)
    eventos.salvar(demo, 1, p[TipoEvento.DIREITO_RECEBIDO].dados, REFS, FISCAIS)
    exercicio = p[TipoEvento.EXERCICIO_DIREITO].dados
    with pytest.raises(ImportacaoError, match="valor pago"):
        eventos.salvar(demo, 1, exercicio, REFS, FISCAIS)
    eventos.salvar(demo, 1, DadosEvento(**{**exercicio.__dict__, "valor": D(1700)}), REFS, FISCAIS)
    antes = _carteira(demo, date(2023, 6, 15))
    assert antes["GGBR10"] == (85, D("1700.00"))
    assert antes["GGBR2"] == (15, D("0.00"))  # os 15 que sobraram, sem custo
    assert antes["GGBR4"][0] == 300
    custo_antes = antes["GGBR4"][1]
    eventos.salvar(demo, 1, p[TipoEvento.INCORPORACAO].dados, REFS, FISCAIS)
    depois = _carteira(demo, date(2023, 6, 16))
    assert "GGBR10" not in depois
    assert depois["GGBR4"] == (385, custo_antes + D(1700))
    assert {x.situacao for x in eventos.propostas(demo, 1, REFS)} == {"lançado"}


def test_subscricao_lancada_do_jeito_antigo_nao_volta_como_pendente(demo: Path) -> None:
    """Quem já tinha lançado a subscrição direto na ação (até a 3.1.1) não vê o exercício como
    pendência nova: lançar de novo dobraria as ações."""
    antiga = DadosEvento(tipo=TipoEvento.SUBSCRICAO, data=date(2023, 5, 2), conta=AGORA_CONTA,
                         ticker="GGBR4", cv="C", quantidade=D(85), valor=D(1700))  # fmt: skip
    eventos.salvar(demo, 1, antiga, REFS, FISCAIS)
    assert _por_tipo(demo)[TipoEvento.EXERCICIO_DIREITO].situacao == "lançado"


def test_mapa_das_linhas_da_subscricao() -> None:
    def mapa(movimentacao: str, entrada: bool, codigo: str) -> object:
        return _mapa_proposta(MovimentacaoB3(movimentacao=movimentacao, entrada=entrada), codigo,
                              FISCAIS)  # fmt: skip

    assert mapa("Direito de Subscrição", True, "XPTO1") == (TipoEvento.DIREITO_RECEBIDO, "C")
    assert mapa("Direito de Subscrição", False, "XPTO1") is None
    assert mapa("Recibo de Subscrição", True, "XPTO9") == (TipoEvento.EXERCICIO_DIREITO, None)
    assert mapa("Recibo de Subscrição", False, "XPTO9") is None
    # recibo de cota de fundo (final 13, 14...): sem o direito deduzido, fica a subscrição simples
    assert mapa("Recibo de Subscrição", True, "XPTO13") == (TipoEvento.SUBSCRICAO, "C")
    assert mapa("Solicitação de Subscrição", False, "XPTO1") is None
    assert mapa("Direitos de Subscrição - Exercido", False, "XPTO1") is None
    assert mapa("Cessão de Direitos", True, "XPTO1") is None
    assert mapa("Atualização", True, "XPTO3") is None  # só com o recibo que a originou


def _banco_com(tmp_path: Path, linhas: list[list[object]]) -> Path:
    banco = criar(tmp_path / "demo")
    arquivo = planilha(
        banco.parent / "movimentacao-2023.xlsx", {"Movimentação": [CAB_MOV, *linhas]}
    )
    importar_planilhas(banco, [arquivo], 1, AGORA)
    return banco


def test_credito_do_direito_dias_antes_nao_esconde_o_exercicio(tmp_path: Path) -> None:
    """Crédito em 24/04 e recibo em 02/05 (8 dias): lançar o crédito não pode marcar o exercício
    como lançado; o exercício lançado não marca a conversão de 09/05 (7 dias depois)."""
    banco = _banco_com(tmp_path, [
        ["Credito", "24/04/2023", "Direito de Subscrição", DIREITO, AGORA_B3, 100, "-", "-"],
        ["Credito", "02/05/2023", "Recibo de Subscrição", RECIBO, AGORA_B3, 85, "-", "-"],
        ["Credito", "09/05/2023", "Atualização", ACAO, AGORA_B3, 85, "-", "-"],
    ])  # fmt: skip
    p = {x.dados.tipo: x for x in eventos.propostas(banco, 1, REFS)}
    eventos.salvar(banco, 1, p[TipoEvento.DIREITO_RECEBIDO].dados, REFS, FISCAIS)
    p = {x.dados.tipo: x for x in eventos.propostas(banco, 1, REFS)}
    assert p[TipoEvento.DIREITO_RECEBIDO].situacao == "lançado"
    assert p[TipoEvento.EXERCICIO_DIREITO].situacao == "pendente"
    exercicio = p[TipoEvento.EXERCICIO_DIREITO].dados
    eventos.salvar(banco, 1, DadosEvento(**{**exercicio.__dict__, "valor": D(1700)}), REFS, FISCAIS)
    p = {x.dados.tipo: x for x in eventos.propostas(banco, 1, REFS)}
    assert p[TipoEvento.EXERCICIO_DIREITO].situacao == "lançado"
    assert p[TipoEvento.INCORPORACAO].situacao == "pendente"
    eventos.salvar(banco, 1, p[TipoEvento.INCORPORACAO].dados, REFS, FISCAIS)
    assert {x.situacao for x in eventos.propostas(banco, 1, REFS)} == {"lançado"}


def test_recibo_vendido_em_parte_converte_so_o_que_sobrou(tmp_path: Path) -> None:
    """Recibo de 85 e ações creditadas de 60 (25 recibos vendidos antes): a conversão proposta é
    de 60, a quantidade das ações. Quantidade maior que a do recibo não é conversão."""
    banco = _banco_com(tmp_path, [
        ["Credito", "02/05/2023", "Recibo de Subscrição", RECIBO, AGORA_B3, 85, "-", "-"],
        ["Credito", "16/06/2023", "Atualização", ACAO, AGORA_B3, 60, "-", "-"],
        ["Credito", "20/06/2023", "Atualização", ACAO, AGORA_B3, 90, "-", "-"],
    ])  # fmt: skip
    lista = eventos.propostas(banco, 1, REFS)
    [conversao] = [x.dados for x in lista if x.dados.tipo is TipoEvento.INCORPORACAO]
    assert (conversao.data, conversao.quantidade, conversao.quantidade_destino) == (
        date(2023, 6, 16), 60, 60)  # fmt: skip
