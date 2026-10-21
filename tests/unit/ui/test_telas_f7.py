"""F7a: telas 8 (opções), 9 (futuros), 10 (proventos), 11 (outros custos), 15 (conciliação) e a
importação das planilhas da B3, sem navegador e sem dados reais."""

import asyncio
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
from nicegui import ui
from nicegui.elements.upload_files import SmallFileUpload
from nicegui.testing.user_simulation import user_simulation

from coir.domain.fiscal import Categoria, TipoEvento
from coir.servicos import eventos, outros_custos
from coir.servicos.b3 import importar_planilhas
from coir.servicos.cadastros import adicionar_titular
from coir.ui.app import montar
from coir.ui.telas.opcoes import do_ano
from tests.apoio import cpf_de_teste
from tests.unit.test_importador_b3 import CAB_ACOES, CAB_MOV, CAB_OPCOES, movimentacao, planilha
from tests.unit.ui.test_telas import AGORA, ESPERA, _estado
from tools.banco_demo import criar


@pytest.fixture(scope="module")
def demo(tmp_path_factory: pytest.TempPathFactory) -> Path:
    return criar(tmp_path_factory.mktemp("demo-f7"))


def _grade(user: object, marca: str) -> ui.aggrid:
    grade = user.find(marker=marca).elements.pop()  # type: ignore[attr-defined]
    assert isinstance(grade, ui.aggrid)
    return grade


@pytest.mark.asyncio
async def test_opcoes_por_situacao_e_trilha_da_serie(demo: Path) -> None:
    """Demonstração: em jan/2023 a BTG abre IRBRE35 e IRBRI258 (abertas no fim do ano) e as
    séries N de fevereiro viram pó em 17/02/2023."""
    estado = _estado(demo)
    estado.titular, estado.ano = 1, 2023
    grupos = do_ano(estado.opcoes(), 2023)
    assert {c.ticker for c in grupos["aberta"]} == {"IRBRE35", "IRBRI258"}
    assert len(grupos["vencida (pó)"]) == 4
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/opcoes")
        await user.should_see("Vencidas (pó) (4)", retries=ESPERA)
        await user.should_see("Abertas (2)")
        user.find(marker="aba-vencida").click()
        grade = _grade(user, "grade-opcoes")
        linhas = grade.options["rowData"]
        assert {r["c0"] for r in linhas} == {c.ticker for c in grupos["vencida (pó)"]}
        assert all(r["c1"].startswith("Put") for r in linhas)  # letra N: put
        user.find(marker="grade-opcoes").trigger("rowClicked", {"data": linhas[0]})
        await user.should_see(f"Trilha de {linhas[0]['c0']}")
        passos = _grade(user, "trilha-opcao").options["rowData"]
        assert "opção não exercida no vencimento (pó)" in passos[-1]["c6"]


@pytest.mark.asyncio
async def test_futuros_resultado_do_ano_e_ajustes_do_contrato(demo: Path) -> None:
    estado = _estado(demo)
    estado.titular, estado.ano = 1, 2024
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/futuros")
        await user.should_see("Resultado de 2024", retries=ESPERA)
        await user.should_see("Nenhum contrato em aberto.")
        grade = _grade(user, "grade-futuros-mes")
        linha = grade.options["rowData"][0]
        assert linha["c0"] == "Janeiro"
        user.find(marker="grade-futuros-mes").trigger("rowClicked", {"data": linha})
        await user.should_see(f"Ajustes de {linha['c1']}")
        ajustes = _grade(user, "grade-ajustes")
        assert ajustes.options["rowData"]


@pytest.mark.asyncio
async def test_proventos_cartoes_grafico_e_tabela(tmp_path: Path) -> None:
    estado = _estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    importar_planilhas(estado.banco, [movimentacao(tmp_path)], 1, AGORA)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/proventos")
        await user.should_see("Dividendos · isentos (09)", retries=ESPERA)
        await user.should_see("100,00")  # 2 x 50,00 de PETR4
        await user.should_see("sem valor")  # JCP da planilha de teste sem valor: aviso
        grade = _grade(user, "grade-proventos")
        assert [r["c1"] for r in grade.options["rowData"]] == ["PETR4", "PETR4"]
        grafico = user.find(marker="grafico-proventos").elements.pop()
        assert isinstance(grafico, ui.echart)


@pytest.mark.asyncio
async def test_outros_custos_grava_o_mes_e_zero_apaga(tmp_path: Path) -> None:
    estado = _estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/outros-custos")
        await user.should_see("Outros custos dedutíveis")
        user.find(marker="descricao-3").type("aluguel de ações")
        user.find(marker="custo-3-COMUM").type("12,50").trigger("blur")
        await user.should_see("Março: reduz o resultado comum em 12,50", retries=ESPERA)
        await user.should_see("Total do ano: 12,50")
        [custo] = outros_custos.listar(estado.banco, 1)
        assert (custo.ano, custo.mes, custo.categoria, custo.valor, custo.descricao) == (
            2026,
            3,
            Categoria.COMUM,
            Decimal("12.50"),
            "aluguel de ações",
        )
        user.find(marker="custo-3-COMUM").clear().type("0").trigger("blur")
        await user.should_see("Total do ano: 0,00", retries=ESPERA)
    assert outros_custos.listar(estado.banco, 1) == []


@pytest.mark.asyncio
async def test_conciliacao_mostra_as_diferencas_da_data(tmp_path: Path) -> None:
    banco = criar(tmp_path / "demo")
    btg = "BANCO BTG PACTUAL S/A."
    consolidado = planilha(tmp_path / "relatorio-consolidado-mensal-2024-janeiro.xlsx", {
        "Posição - Ações": [CAB_ACOES, ["ITSA4 - ITAUSA", btg, "2", "ITSA4", "-", "BRITSAACNPR7",
                                        "PN", "ITAU", 50]],
        "Posição - Opções": [CAB_OPCOES, ["Opção", btg, "Compra", "Opção", "IRBRE35", "E35",
                                          "Comprador", 600]],
    })  # fmt: skip
    importar_planilhas(banco, [consolidado], 1, AGORA)
    estado = _estado(banco)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/conciliacao")
        await user.should_see("Posição em 31/01/2024", retries=ESPERA)
        grade = _grade(user, "grade-conciliacao")
        por_codigo = {r["c1"]: r for r in grade.options["rowData"]}
        assert "IRBRE35" not in por_codigo  # confere: fora do filtro "só as diferenças"
        assert "só na B3" in por_codigo["ITSA4"]["c5"]
        user.find(marker="so-diferencas").click()
        grade = _grade(user, "grade-conciliacao")
        assert any(r["c1"] == "IRBRE35" for r in grade.options["rowData"])


@pytest.mark.asyncio
async def test_importar_planilhas_da_b3(tmp_path: Path) -> None:
    estado = _estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    arquivo = movimentacao(tmp_path)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/importar")
        user.find(marker="aba-b3").click()
        [envio] = user.find(marker="enviar-planilhas").elements
        assert isinstance(envio, ui.upload)
        tipo = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        await envio.handle_uploads([SmallFileUpload(arquivo.name, tipo, arquivo.read_bytes())])
        await user.should_see(arquivo.name, retries=ESPERA)
        await user.should_see("movimentação · 4 nova(s) de 4")
    assert (estado.dados / "_relatorios-b3" / arquivo.name).exists()
    estado.titular = 1
    assert [p.data for p in estado.proventos()[0]] == [date(2026, 9, 15)] * 2


@pytest.mark.asyncio
async def test_eventos_lanca_a_proposta_da_b3_e_recusa_o_que_nao_fecha(tmp_path: Path) -> None:
    """Tela 7: a proposta (desdobro da movimentação) vira evento com a data ex; uma incorporação
    maior que a posição não é gravada e o motivo aparece."""
    banco = criar(tmp_path / "demo")
    agora_b3 = "AGORA CORRETORA DE TITULOS E VALORES MOBILIARIOS S/A"
    movimentos = planilha(tmp_path / "movimentacao-2023.xlsx", {"Movimentação": [
        CAB_MOV,
        ["Credito", "03/03/2023", "Desdobro", "GGBR4 - GERDAU S.A.", agora_b3, 300, "-", "-"],
    ]})  # fmt: skip
    importar_planilhas(banco, [movimentos], 1, AGORA)
    estado = _estado(banco)
    estado.titular, estado.ano = 1, 2023
    from tests.unit.test_servicos_eventos import _pacote_eventos

    estado.pasta_refs = _pacote_eventos(tmp_path, [])  # sem os eventos reais do pacote da B3
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/eventos")
        await user.should_see("Propostas da B3 (1)", retries=ESPERA)
        await user.should_see("Eventos lançados (0)")
        linha = _grade(user, "grade-propostas").options["rowData"][0]
        user.find(marker="grade-propostas").trigger("rowClicked", {"data": linha})
        await user.should_see("Proposta da B3: Desdobro GGBR4")
        [data] = user.find(marker="data-evento").elements
        assert isinstance(data, ui.input)
        data.set_value("2023-03-01")  # data ex (a B3 traz o crédito)
        user.find(marker="gravar-evento").click()
        await user.should_see("Eventos lançados (1)", retries=ESPERA)
        await user.should_see("Propostas da B3 (0)")
        [tipo] = user.find(marker="tipo-evento").elements
        assert isinstance(tipo, ui.select)
        tipo.set_value("INCORPORACAO")
        await user.should_see("Ativo recebido")
        user.find(marker="ativo-evento").type("GGBR4")
        user.find(marker="quantidade-evento").type("1000")  # a posição é 600
        user.find(marker="recebido-evento").type("GOAU4")
        user.find(marker="qtd-recebida-evento").type("500")
        user.find(marker="gravar-evento").click()
        await user.should_see("não fecha com as operações", retries=ESPERA)
        await user.should_see("Eventos lançados (1)")
    [evento] = eventos.listar(banco, 1)
    assert (evento.dados.tipo, evento.dados.data, evento.dados.origem) == (
        TipoEvento.DESDOBRAMENTO,
        date(2023, 3, 1),
        "b3",
    )


@pytest.mark.asyncio
async def test_carteira_exporta_a_posicao_da_data(demo: Path) -> None:
    """A posição exportada tem arquivos próprios (não sobrescreve os relatórios do ano)."""
    estado = _estado(demo)
    estado.titular, estado.ano = 1, 2023
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/carteira")
        await user.should_see("posições em", retries=ESPERA)
        user.find(marker="exportar-carteira").click()
        await user.should_see("exportada", retries=ESPERA)
    pasta = demo.parent / "relatorios" / "carteira"
    nomes = sorted(p.name for p in pasta.iterdir())
    prefixo = "coir-carteira-2023-12-31-por-titular-titular1"
    assert nomes and all(n.startswith(prefixo) for n in nomes)


@pytest.mark.asyncio
async def test_configuracoes_atualiza_as_tabelas_e_o_painel_avisa(
    tmp_path: Path, configuracao_temporaria: Path
) -> None:
    """F7b: com a opção "verificar ao abrir", o Painel avisa do pacote novo; em Configurações,
    verificar mostra o resumo e aplicar passa o pacote (conferido) a ser o em uso."""
    from coir.ui import preferencias
    from tests.unit.test_atualizacao_referencias import _github, _pacote

    estado = _estado(tmp_path / "dados" / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    estado.baixar = _github(_pacote(tmp_path))
    preferencias.escolher_verificar_tabelas(True)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/")
        await user.should_see("Há tabelas de referência novas (01/01/2099).", retries=ESPERA)
        await user.open("/configuracoes")
        await user.should_see("o da instalação")
        user.find(marker="verificar-tabelas").click()
        await user.should_see("Novo: 01/01/2099", retries=ESPERA)
        user.find(marker="aplicar-tabelas").click()
        await user.should_see("Tabelas atualizadas", retries=ESPERA)
        await user.should_see("Em uso: pacote de 01/01/2099 (baixado)")
    assert estado.refs.ativos  # relidas do pacote baixado


@pytest.mark.asyncio
async def test_proventos_data_do_credito_do_jcp_muda_o_ano(tmp_path: Path) -> None:
    """F8: JCP pago em 10/03/2026 a 17,5% (sem data: declarado em 2026); o usuário informa o
    crédito em 15/12/2025 pelo informe e o JCP passa para 2025."""
    from sqlalchemy import select

    from coir.db.modelos import MovimentacaoB3
    from coir.db.sessao import conectar

    estado = _estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    mov = planilha(tmp_path / "movimentacao-2026-10-16.xlsx", {"Movimentação": [
        CAB_MOV,
        ["Credito", "10/03/2026", "Juros Sobre Capital Próprio", "ITSA4 - ITAUSA",
         "XP INVESTIMENTOS CCTVM S/A", 1000, "0,5", "412,50"],
    ]})  # fmt: skip
    importar_planilhas(estado.banco, [mov], 1, AGORA)
    with conectar(estado.banco) as fabrica, fabrica() as s:
        chave = s.scalars(select(MovimentacaoB3.chave)).one()
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/proventos")
        await user.should_see("Data do crédito do JCP", retries=ESPERA)
        await user.should_see("declarado em 2026 · sem data")
        user.find(marker=f"credito-{chave[:12]}").type("15/12/2025")
        user.find(marker=f"salvar-credito-{chave[:12]}").click()
        await user.should_see("declarado em 2025 · informada", retries=ESPERA)


@pytest.mark.asyncio
async def test_declaracao_anual_ficha_a_ficha_e_lancado(demo: Path) -> None:
    """F8, tela 14: na demonstração, 2025 tem mapeamento (exercício 2026): Bens e Direitos com o
    item em carteira, renda variável por mês e a marcação "lançado" gravada no banco; 2023 não
    tem mapeamento e a tela avisa."""
    estado = _estado(demo)
    estado.titular, estado.ano = 1, 2025
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/declaracao")
        await user.should_see("Exercício 2026, ano-calendário 2025", retries=ESPERA)
        await user.should_see("03.01", retries=ESPERA)
        user.find(marker="ficha-rv").click()
        await user.should_see("Mercado à Vista - Ações", retries=ESPERA)
        user.find(marker="lancado-rv:01").click()
        for _ in range(ESPERA):  # a gravação roda numa thread
            if "rv:01" in estado.lancados():
                break
            await asyncio.sleep(0.1)
        assert "rv:01" in estado.lancados()
        user.find(marker="ficha-conferencia").click()
        await user.should_see("Nenhuma DIRPF deste ano-calendário", retries=ESPERA)
    estado.marcar_lancado("rv:01", False)
    estado.ano = 2023
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/declaracao")
        await user.should_see("Sem o mapeamento das fichas do programa IRPF 2024",
                              retries=ESPERA)  # fmt: skip
