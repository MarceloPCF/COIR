"""Tela 7 em abas: formulários curtos por tarefa e o botão "Lançar as seguras"; botão "Como faço?"
de cada tela. Sem navegador e sem dados reais."""

import asyncio
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Any

import pytest
from nicegui import ui
from nicegui.testing import User
from nicegui.testing.user_simulation import user_simulation

from coir.domain.fiscal import NaturezaEvento, TipoEvento, Visao
from coir.referencias import carregar, carregar_fiscais
from coir.servicos import eventos
from coir.servicos.apuracao import carteira_banco
from coir.ui.ajuda import arquivo_da_tela, texto_da_tela
from coir.ui.app import INICIO, TELAS, montar
from tests.unit.test_servicos_eventos import _pacote_eventos
from tests.unit.ui.test_telas import ESPERA, _estado
from tools.banco_demo import criar

REFS, FISCAIS = carregar(), carregar_fiscais()
D = Decimal
AGORA_CONTA = 1  # GGBR4: 300 ações na demonstração
BTG_CONTA = 2


def _campo(user: User, marca: str) -> ui.input:
    [campo] = user.find(marker=marca).elements
    assert isinstance(campo, ui.input)
    return campo


def _seletor(user: User, marca: str) -> ui.select:
    [campo] = user.find(marker=marca).elements
    assert isinstance(campo, ui.select)
    return campo


def _linha(user: User, marca: str) -> dict[str, Any]:
    [grade] = user.find(marker=marca).elements
    assert isinstance(grade, ui.aggrid)
    return grade.options["rowData"][0]


def _posicao(banco: Path, ticker: str, dia: date) -> tuple[Decimal, Decimal]:
    posicoes, _ = carteira_banco(banco, REFS, FISCAIS, dia, Visao.TITULAR)
    [p] = [x for x in posicoes if x.ticker == ticker]
    return p.quantidade, p.custo


@pytest.fixture
def demo(tmp_path: Path) -> Path:
    return criar(tmp_path / "demo")


@pytest.mark.asyncio
async def test_oferta_publica_abre_posicao_de_ativo_sem_nota(demo: Path, tmp_path: Path) -> None:
    """A aba Ofertas pede só data, conta, ativo, ações e valor; o ativo não precisa ter nota."""
    estado = _estado(demo)
    estado.titular, estado.ano = 1, 2023
    estado.pasta_refs = _pacote_eventos(tmp_path, [])
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/eventos")
        await user.should_see("Ofertas públicas (0)", retries=ESPERA)
        user.find(marker="aba-ofertas").click()
        _campo(user, "ofertas-data-evento").set_value("2023-06-15")
        user.find(marker="ofertas-ativo-evento").type("NOVA3")
        user.find(marker="ofertas-quantidade-evento").type("100")
        user.find(marker="ofertas-valor-evento").type("2.500,00")
        user.find(marker="ofertas-gravar-evento").click()
        await user.should_see("Ofertas públicas (1)", retries=ESPERA)
        assert _linha(user, "grade-ofertas")["c1"] == "NOVA3"
    [evento] = eventos.listar(demo, 1)
    assert (evento.dados.tipo, evento.dados.cv) == (TipoEvento.OFERTA_PUBLICA, "C")
    assert _posicao(demo, "NOVA3", date(2023, 12, 31)) == (D(100), D(2500))


@pytest.mark.asyncio
async def test_transferencia_sem_cv_e_com_toda_a_posicao(demo: Path, tmp_path: Path) -> None:
    """Quantidade vazia move toda a posição; a lista mostra origem, destino e "toda a posição"."""
    estado = _estado(demo)
    estado.titular, estado.ano = 1, 2023
    estado.pasta_refs = _pacote_eventos(tmp_path, [])
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/eventos")
        await user.should_see("Transferências (0)", retries=ESPERA)
        user.find(marker="aba-transferencias").click()
        _campo(user, "transferencias-data-evento").set_value("2023-06-15")
        _seletor(user, "transferencias-conta-evento").set_value(AGORA_CONTA)
        _seletor(user, "transferencias-destino-evento").set_value(BTG_CONTA)
        user.find(marker="transferencias-ativo-evento").type("GGBR4")
        user.find(marker="transferencias-gravar-evento").click()
        await user.should_see("Transferências (1)", retries=ESPERA)
        await user.should_see("Eventos lançados (1)")
        assert _linha(user, "grade-transferencias")["c4"] == "toda a posição"
    [evento] = eventos.listar(demo, 1)
    assert (evento.dados.tipo, evento.dados.cv) == (TipoEvento.TRANSFERENCIA, None)


@pytest.mark.asyncio
async def test_evento_seguro_entra_sozinho_e_da_para_desfazer(demo: Path, tmp_path: Path) -> None:
    """Ao abrir a tela, o desdobramento entra sozinho e dá para desfazer; a bonificação sem custo
    fica para conferir. Desfeito, ele fica ignorado e não volta ao abrir a tela de novo."""
    estado = _estado(demo)
    estado.titular, estado.ano = 1, 2023
    estado.pasta_refs = _pacote_eventos(tmp_path, [
        "GGBR,BRGGBRTESTE1,DESDOBRAMENTO,100,2023-02-28,2023-02-01,BRGGBRTESTE1,,,",
        "GGBR,BRGGBRTESTE1,BONIFICACAO,5,2023-05-31,2023-05-01,BRGGBRTESTE1,,,",
    ])  # fmt: skip
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/eventos")
        await user.should_see("1 evento(s) lançado(s)", retries=ESPERA)
        await user.should_see("Eventos lançados (1)", retries=ESPERA)
        await user.should_see("A conferir (1)", retries=ESPERA)
        assert len(eventos.listar(demo, 1)) == 1
        user.find(marker="desfazer-lote").click()
        await user.should_see("Lançamento desfeito.", retries=ESPERA)
        await user.should_see("Eventos lançados (0)", retries=ESPERA)
    assert eventos.listar(demo, 1) == []
    assert estado.lancar_eventos_automaticos().ids == ()
    assert eventos.listar(demo, 1) == []


@pytest.mark.asyncio
async def test_como_faco_abre_a_pagina_do_manual_sem_captura(demo: Path) -> None:
    estado = _estado(demo)
    estado.titular, estado.ano = 1, 2023
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/darf")
        user.find(marker="como-faco").click()
        [titulo] = user.find(marker="ajuda-titulo").elements
        assert isinstance(titulo, ui.label)
        for _ in range(ESPERA):  # o clique roda numa tarefa
            if titulo.text:
                break
            await asyncio.sleep(0.1)
        assert titulo.text == "DARF"
        [texto] = user.find(marker="ajuda-texto").elements
        assert isinstance(texto, ui.markdown)
        assert "6015" in texto.content and "imagens" not in texto.content


@pytest.mark.parametrize("rota", [t.rota for t in TELAS])
def test_toda_tela_tem_pagina_no_manual(rota: str) -> None:
    titulo, corpo = texto_da_tela(arquivo_da_tela(rota).read_text(encoding="utf-8"))
    assert titulo and corpo
    assert "](" not in corpo.replace("](http", "")  # sem links relativos, que só valem no site
    assert INICIO.rota not in {t.rota for t in TELAS}


def test_o_executavel_leva_o_manual() -> None:
    spec = Path(__file__).resolve().parents[3] / "empacotamento" / "coir.spec"
    assert '"docs" / "manual" / "telas" / "*.md"' in spec.read_text(encoding="utf-8")


@pytest.mark.asyncio
async def test_outros_custos_avisa_quando_a_descricao_fala_em_aluguel(tmp_path: Path) -> None:
    """A tela antiga mandava lançar a taxa de aluguel: quem seguiu o texto precisa revisar."""
    from coir.domain.fiscal import Categoria
    from coir.servicos import outros_custos
    from coir.servicos.cadastros import adicionar_titular
    from tests.apoio import cpf_de_teste
    from tests.unit.ui.test_telas import AGORA

    estado = _estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    outros_custos.gravar(estado.banco, 1, 2026, 3, Categoria.COMUM, D("10"), "aluguel de ações")
    estado.titular, estado.ano = 1, 2026
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/outros-custos")
        await user.should_see("Não lance a taxa de aluguel", retries=ESPERA)
        await user.should_see("Há lançamento com aluguel na descrição")


@pytest.mark.asyncio
async def test_bonificacao_gravada_sem_custo_e_avisada_e_abre_para_corrigir(
    demo: Path, tmp_path: Path
) -> None:
    """Caso UNIP6 (bonificação lançada sem valor): a aba A conferir avisa e o botão Corrigir abre
    o evento em Lançados, para informar o custo atribuído."""
    eventos.salvar(demo, 1, eventos.DadosEvento(
        tipo=TipoEvento.BONIFICACAO, data=date(2023, 6, 1), conta=AGORA_CONTA, ticker="GGBR4",
        cv="C", quantidade=D(30), valor=D(0)), REFS, FISCAIS)  # fmt: skip
    estado = _estado(demo)
    estado.titular, estado.ano = 1, 2023
    estado.pasta_refs = _pacote_eventos(tmp_path, [])
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/eventos")
        await user.should_see(marker="bonificacao-sem-custo", retries=ESPERA)
        await user.should_see("lançada sem custo atribuído", retries=ESPERA)
        user.find(marker="corrigir-sem-custo").click()
        await user.should_see("Corrigir: Bonificação", retries=ESPERA)


@pytest.mark.asyncio
async def test_aba_subscricoes_direito_recebido_exercicio_e_recibo(
    demo: Path, tmp_path: Path
) -> None:
    """Os três lançamentos pela tela: 100 direitos GGBR2 recebidos em 03/04/2023; exercício de 85
    em 02/05 pagando 1.700,00 (recibo GGBR10 com 85 por 1.700,00); em 16/06 o recibo vira GGBR4,
    que passa de 300 a 385 ações com 1.700,00 a mais de custo."""
    estado = _estado(demo)
    estado.titular, estado.ano = 1, 2023
    estado.pasta_refs = _pacote_eventos(tmp_path, [])
    _, custo_antes = _posicao(demo, "GGBR4", date(2023, 4, 1))
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/eventos")
        await user.should_see("Subscrições (0)", retries=ESPERA)
        user.find(marker="aba-subscricoes").click()
        # o formulário abre no exercício; troca para o direito recebido
        _seletor(user, "subscricoes-tipo-evento").set_value(TipoEvento.DIREITO_RECEBIDO.value)
        await user.should_see("Direitos recebidos", retries=ESPERA)
        await user.should_not_see(marker="subscricoes-valor-evento")  # entra com custo zero
        _campo(user, "subscricoes-data-evento").set_value("2023-04-03")
        _seletor(user, "subscricoes-conta-evento").set_value(AGORA_CONTA)
        user.find(marker="subscricoes-ativo-evento").type("GGBR2")
        user.find(marker="subscricoes-quantidade-evento").type("100")
        user.find(marker="subscricoes-gravar-evento").click()
        await user.should_see("Subscrições (1)", retries=ESPERA)
        # volta ao exercício: direitos usados, recibo recebido e valor pago
        await user.should_see("Direitos exercidos", retries=ESPERA)
        _campo(user, "subscricoes-data-evento").set_value("2023-05-02")
        _seletor(user, "subscricoes-conta-evento").set_value(AGORA_CONTA)
        user.find(marker="subscricoes-ativo-evento").type("GGBR2")
        user.find(marker="subscricoes-quantidade-evento").type("85")
        user.find(marker="subscricoes-valor-evento").type("1.700,00")
        user.find(marker="subscricoes-recebido-evento").type("GGBR10")
        user.find(marker="subscricoes-qtd-recebida-evento").type("85")
        user.find(marker="subscricoes-gravar-evento").click()
        await user.should_see("Subscrições (2)", retries=ESPERA)
    assert _posicao(demo, "GGBR10", date(2023, 5, 31)) == (D(85), D(1700))
    assert _posicao(demo, "GGBR2", date(2023, 5, 31)) == (D(15), D(0))
    recibo = eventos.DadosEvento(
        tipo=TipoEvento.INCORPORACAO, data=date(2023, 6, 16), conta=AGORA_CONTA, ticker="GGBR10",
        quantidade=D(85), ticker_destino="GGBR4", quantidade_destino=D(85),
        natureza=NaturezaEvento.RECIBO_SUBSCRICAO,
    )  # fmt: skip
    eventos.salvar(demo, 1, recibo, REFS, FISCAIS)
    assert _posicao(demo, "GGBR4", date(2023, 6, 30)) == (D(385), custo_antes + D(1700))
    tipos = [e.dados.tipo for e in eventos.listar(demo, 1)]
    assert set(tipos) == {TipoEvento.DIREITO_RECEBIDO, TipoEvento.EXERCICIO_DIREITO,
                          TipoEvento.INCORPORACAO}  # fmt: skip


@pytest.mark.asyncio
async def test_dividendo_pago_em_bens_pela_aba_lancados(demo: Path, tmp_path: Path) -> None:
    """B10(d) do parecer de 16/11/2026: a GGBR4 paga dividendo com 20 cotas de KNRI11, valor
    atribuído 2.000,00 e valor de mercado na entrega 1.900,00: as cotas entram por 1.900,00."""
    estado = _estado(demo)
    estado.titular, estado.ano = 1, 2023
    estado.pasta_refs = _pacote_eventos(tmp_path, [])
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/eventos")
        await user.should_see("Transferências (0)", retries=ESPERA)
        user.find(marker="aba-lancados").click()
        _seletor(user, "lancados-tipo-evento").set_value(TipoEvento.DIVIDENDO_EM_BENS.value)
        await user.should_see("Empresa que pagou (código)", retries=ESPERA)
        await user.should_not_see(marker="lancados-quantidade-evento")
        _campo(user, "lancados-data-evento").set_value("2023-03-01")
        _seletor(user, "lancados-conta-evento").set_value(AGORA_CONTA)
        user.find(marker="lancados-ativo-evento").type("GGBR4")
        user.find(marker="lancados-valor-evento").type("2.000,00")
        user.find(marker="lancados-valor-corrente-evento").type("1.900,00")
        user.find(marker="lancados-recebido-evento").type("KNRI11")
        user.find(marker="lancados-qtd-recebida-evento").type("20")
        user.find(marker="lancados-gravar-evento").click()
        await user.should_see("Eventos lançados (1)", retries=ESPERA)
    assert _posicao(demo, "KNRI11", date(2023, 3, 31)) == (D(20), D(1900))
