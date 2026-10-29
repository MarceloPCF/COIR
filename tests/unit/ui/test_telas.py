"""F6: fluxos das telas 1, 3, 4 e 5 sem navegador (assistente, importar, notas, operações)."""

import asyncio
import json
from datetime import datetime
from pathlib import Path

import pytest
from nicegui import ui
from nicegui.elements.upload_files import SmallFileUpload
from nicegui.testing.user_simulation import user_simulation

from coir.servicos import manutencao
from coir.servicos.cadastros import adicionar_titular
from coir.servicos.estado import Estado
from coir.servicos.notas import listar
from coir.ui.app import montar
from tests.apoio import cpf_de_teste
from tools import notas_ficticias
from tools.banco_demo import criar, extrator, notas

AGORA = datetime(2026, 10, 16, 12, 0, 0)
ESPERA = 80  # tentativas de 0,1 s (a importação roda numa thread)


def _estado(banco: Path) -> Estado:
    estado = Estado(banco)
    estado.agora = lambda: AGORA
    estado.extrator = extrator(cpf_de_teste())
    return estado


@pytest.fixture(scope="module")
def demo(tmp_path_factory: pytest.TempPathFactory) -> Path:
    return criar(tmp_path_factory.mktemp("demo"))


@pytest.mark.asyncio
async def test_assistente_do_zero_pasta_titular_e_notas(
    tmp_path: Path, configuracao_temporaria: Path
) -> None:
    estado = _estado(tmp_path / "inicial" / "coir.db")
    dados = tmp_path / "meus-dados"
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/")
        await user.should_see("Bem-vindo ao COIR")
        user.find(marker="pasta-dados").clear().type(str(dados))
        user.find(marker="continuar-pasta").click()
        await user.should_see("Quem é o titular das notas?")
        assert estado.banco == dados / "coir.db" and estado.existe
        assert json.loads(configuracao_temporaria.read_text(encoding="utf-8")) == {
            "dados": str(dados)
        }
        user.find(marker="nome").type("Titular Novo")
        user.find(marker="cpf").type(cpf_de_teste())
        user.find(marker="cadastrar").click()
        await user.should_see("Importe as notas de corretagem")
        await user.should_see("Titular Novo")  # já no seletor do topo
        assert estado.pronto
        user.find(marker="concluir").click()
        await user.should_see("Painel")


@pytest.mark.asyncio
async def test_assistente_recusa_cpf_invalido(tmp_path: Path) -> None:
    estado = _estado(tmp_path / "coir.db")
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/")
        user.find(marker="continuar-pasta").click()
        await user.should_see("Quem é o titular das notas?")
        user.find(marker="nome").type("Fulano")
        user.find(marker="cpf").type("11111111111")
        user.find(marker="cadastrar").click()
        await user.should_see("CPF inválido")
    assert estado.titulares() == []


@pytest.mark.asyncio
async def test_importar_pdfs_escolhidos_copia_para_a_pasta_e_mostra_cada_arquivo(
    tmp_path: Path,
) -> None:
    estado = _estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    arquivos = notas()[:2]
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/importar")
        await user.should_see("Notas de corretagem (PDF)")
        [envio] = user.find(marker="enviar-pdfs").elements
        assert isinstance(envio, ui.upload)
        await envio.handle_uploads(
            [SmallFileUpload(f"{a.stem}.pdf", "application/pdf", a.read_bytes()) for a in arquivos]
        )
        await user.should_see("notas gravadas", retries=ESPERA)
        for a in arquivos:
            await user.should_see(f"{a.stem}.pdf")
    copiados = sorted(p.name for p in estado.pasta_notas.iterdir())
    assert copiados == sorted(f"{a.stem}.pdf" for a in arquivos)
    assert len(listar(estado.banco, None, None)) >= len(arquivos)


@pytest.mark.asyncio
async def test_importar_a_pasta_das_notas(tmp_path: Path) -> None:
    estado = _estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    estado.pasta_notas.mkdir()
    origem = notas()[0]
    (estado.pasta_notas / f"{origem.stem}.pdf").write_bytes(origem.read_bytes())
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/importar")
        user.find(marker="importar-pasta").click()
        await user.should_see("notas gravadas", retries=ESPERA)
        await user.should_see("gravado")
    assert listar(estado.banco, None, None)
    assert [p.name for p in estado.pasta_notas.iterdir()] == [f"{origem.stem}.pdf"]  # sem cópia


@pytest.mark.asyncio
async def test_notas_abre_a_nota_com_a_conciliacao(demo: Path) -> None:
    estado = _estado(demo)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/notas")
        await user.should_see("Escolha uma nota na lista.")
        lista = user.find(marker="lista-notas").elements.pop()
        assert isinstance(lista, ui.aggrid)
        primeira = lista.options["rowData"][0]
        user.find(marker="lista-notas").trigger("rowClicked", {"data": primeira})
        await user.should_see("Conciliada", retries=ESPERA)
        await user.should_see("não está na pasta das notas")  # fixtures: sem PDF


@pytest.mark.asyncio
async def test_notas_amplia_a_pagina_do_pdf(tmp_path: Path) -> None:
    """A página ampliada aparecia em branco (q-img preguiçoso dentro do diálogo): agora é uma
    imagem simples, renderizada maior ao abrir."""
    estado = _estado(criar(tmp_path / "demo", pdfs=True))
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/notas")
        lista = user.find(marker="lista-notas").elements.pop()
        assert isinstance(lista, ui.aggrid)
        user.find(marker="lista-notas").trigger("rowClicked", {"data": lista.options["rowData"][0]})
        await user.should_see(marker="pagina-pdf", retries=ESPERA)
        user.find(marker="pagina-pdf").trigger("click")
        ampliada = user.find(marker="pagina-ampliada").elements.pop()
        assert isinstance(ampliada, ui.interactive_image)
        for _ in range(ESPERA):
            if ampliada.source:
                break
            await asyncio.sleep(0.1)
        assert str(ampliada.source).startswith("data:image/png;base64,")


@pytest.mark.asyncio
async def test_notas_filtra_pelo_numero(demo: Path) -> None:
    estado = _estado(demo)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/notas")
        lista = user.find(marker="lista-notas").elements.pop()
        assert isinstance(lista, ui.aggrid)
        total = len(lista.options["rowData"])
        numero = lista.options["rowData"][0]["c2"]
        user.find(marker="filtro-numero").type(numero)
        await user.should_see(f"de {total} nota")
        assert all(numero in r["c2"] for r in lista.options["rowData"])


@pytest.mark.asyncio
async def test_operacoes_filtra_por_ativo_e_soma_no_rodape(demo: Path) -> None:
    estado = _estado(demo)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/operacoes")
        grade = user.find(marker="grade-operacoes").elements.pop()
        assert isinstance(grade, ui.aggrid)
        linhas = grade.options["rowData"]
        assert linhas and grade.options["pinnedBottomRowData"][0]["c0"] == (
            f"{len(linhas)} operações"
        )
        ativo = linhas[0]["c4"]
        user.find(marker="filtro-ativo").type(ativo)
        await user.should_see("Operações")
        assert {r["c4"] for r in grade.options["rowData"]} <= {
            r["c4"] for r in linhas if ativo in r["c4"]
        }


@pytest.mark.asyncio
async def test_operacoes_historico_de_todos_os_anos(demo: Path) -> None:
    """O histórico de um ativo em todos os anos: com o interruptor, a grade passa a ter as
    operações de todos os anos (a demonstração tem notas de 2008 a 2026)."""
    estado = _estado(demo)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/operacoes")
        grade = user.find(marker="grade-operacoes").elements.pop()
        assert isinstance(grade, ui.aggrid)
        do_ano = {r["c0"][-4:] for r in grade.options["rowData"]}
        chave = user.find(marker="filtro-todos-os-anos").elements.pop()
        assert isinstance(chave, ui.switch)
        chave.set_value(True)
        await user.should_see("Operações")
        anos = {r["c0"][-4:] for r in grade.options["rowData"]}
    assert do_ano == {"2026"} and len(anos) > 1


@pytest.mark.asyncio
async def test_clicar_numa_operacao_abre_a_nota(demo: Path) -> None:
    estado = _estado(demo)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/operacoes")
        grade = user.find(marker="grade-operacoes").elements.pop()
        assert isinstance(grade, ui.aggrid)
        linha = grade.options["rowData"][0]
        user.find(marker="grade-operacoes").trigger("rowClicked", {"data": linha})
        await user.should_see(f"Nota {linha['c3']}", retries=ESPERA)
        await user.should_see("Conciliada", retries=ESPERA)


@pytest.mark.asyncio
async def test_endereco_de_nota_invalido(demo: Path) -> None:
    estado = _estado(demo)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/notas/abc")
        await user.should_see("Nota 'abc' inválida.")


@pytest.mark.asyncio
async def test_carteira_abre_a_trilha_do_pm(demo: Path) -> None:
    estado = _estado(demo)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/carteira")
        await user.should_see("posições em", retries=ESPERA)
        grade = user.find(marker="grade-carteira").elements.pop()
        assert isinstance(grade, ui.aggrid)
        linha = grade.options["rowData"][0]
        user.find(marker="grade-carteira").trigger("rowClicked", {"data": linha})
        await user.should_see(f"Trilha do PM: {linha['c0']}", retries=ESPERA)


@pytest.mark.asyncio
async def test_apuracao_abre_a_memoria_do_mes(demo: Path) -> None:
    estado = _estado(demo)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/apuracao")
        await user.should_see("Apurado até 16/10/2026", retries=ESPERA)
        grade = user.find(marker="grade-apuracao").elements.pop()
        assert isinstance(grade, ui.aggrid)
        agosto = next(r for r in grade.options["rowData"] if r["c0"] == "Agosto")
        user.find(marker="grade-apuracao").trigger("rowClicked", {"data": agosto})
        await user.should_see("Memória de agosto de 2026")
        await user.should_see("1. Limite de isenção (ações à vista)")
        await user.should_see("Movimentos do mês com resultado (13)")


@pytest.mark.asyncio
async def test_relatorios_mostra_e_exporta(demo: Path, tmp_path: Path) -> None:
    estado = _estado(demo)
    destino = tmp_path / "relatorios"
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/relatorios")
        await user.should_see("Resumo fiscal do ano", retries=ESPERA)
        user.find(marker="relatorio-livro-de-notas").click()
        await user.should_see("Livro de notas", retries=ESPERA)
        user.find(marker="formato-pdf").click()  # desmarca
        user.find(marker="formato-csv").click()
        user.find(marker="pasta-relatorios").clear().type(str(destino))
        user.find(marker="exportar").click()
        await user.should_see("arquivo(s) gravado(s)", retries=ESPERA)
    assert [p.name for p in destino.iterdir()] == ["coir-2026-titular1.xlsx"]


@pytest.mark.asyncio
async def test_relatorio_com_muitas_secoes_mostra_uma_por_vez(
    demo: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A conferência do PM tem uma seção por ativo: dezenas de grades de uma vez derrubavam a
    conexão da página; acima do limite, uma grade por vez, escolhida na lista."""
    from coir.ui.telas import relatorios as tela

    monkeypatch.setattr(tela, "MAX_SECOES", 1)
    estado = _estado(demo)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/relatorios")
        await user.should_see("Resumo fiscal do ano", retries=ESPERA)
        user.find(marker="relatorio-conferencia-pm").click()
        await user.should_see(marker="secao-relatorio", retries=ESPERA)
        escolha = user.find(marker="secao-relatorio").elements.pop()
        assert isinstance(escolha, ui.select) and len(escolha.options) > 1


@pytest.mark.asyncio
async def test_operacoes_finalizadas_numa_grade_com_filtros_e_resumo(demo: Path) -> None:
    """Operações finalizadas do ano numa grade só, com o resumo por categoria e tipo; filtrar
    pelo ativo deixa só as dele e o resumo acompanha (2016: day trade da nota XP antiga)."""
    estado = _estado(demo)
    estado.ano = 2016
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/relatorios")
        await user.should_see("Resumo fiscal do ano", retries=ESPERA)
        user.find(marker="relatorio-operacoes-finalizadas").click()
        await user.should_see("Resultado por categoria e tipo", retries=ESPERA)
        grade = user.find(marker="finalizadas").elements.pop()
        resumo = user.find(marker="finalizadas-resumo").elements.pop()
        assert isinstance(grade, ui.aggrid) and isinstance(resumo, ui.aggrid)
        ativo = grade.options["rowData"][0]["c1"]
        user.find(marker="finalizadas-ativo").type(ativo)
        await user.should_see("Resultado por categoria e tipo")
        assert {r["c1"] for r in grade.options["rowData"]} == {ativo}
        assert resumo.options["pinnedBottomRowData"][0]["c2"] == str(len(grade.options["rowData"]))


@pytest.mark.asyncio
async def test_darf_marcar_como_paga(tmp_path: Path) -> None:
    from datetime import date
    from decimal import Decimal

    from coir.servicos import darf
    from tests.unit.test_servicos_darf import MESES

    estado = _estado(tmp_path / "coir.db")
    estado.agora = lambda: datetime(2026, 4, 2, 10, 0, 0)
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    estado.darfs = lambda: darf.darfs(  # type: ignore[method-assign]
        MESES, darf.pagamentos(estado.banco, 1), date(2026, 4, 2), estado.fiscais
    )
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/darf")
        await user.should_see("2 DARF(s) em atraso", retries=ESPERA)  # jan e fev
        grade = user.find(marker="grade-darf").elements.pop()
        assert isinstance(grade, ui.aggrid)
        fevereiro = next(r for r in grade.options["rowData"] if r["c0"] == "02/2026")
        assert fevereiro["c5"] == "em atraso"
        user.find(marker="grade-darf").trigger("rowClicked", {"data": fevereiro})
        await user.should_see("DARF de fevereiro/2026")
        await user.should_see("93,10")  # 91,59 + multa de 0,60 + juros de 0,91 até hoje
        user.find(marker="data-pagamento").clear().type("2026-04-01")
        user.find(marker="valor-pagamento").clear().type("93,10")
        user.find(marker="salvar-pagamento").click()
        await user.should_see("Pagamento registrado", retries=ESPERA)
        await user.should_see("Pago em 01/04/2026: 93,10", retries=ESPERA)
    pago = darf.pagamentos(estado.banco, 1)[(2026, 2)]
    assert (pago.data, pago.valor) == (date(2026, 4, 1), Decimal("93.10"))


@pytest.mark.asyncio
async def test_painel_abre_a_darf_atrasada_de_outro_ano(demo: Path) -> None:
    """A DARF da vez é de 2023 com o Painel em 2026: o botão leva à DARF no ano dela, já aberta,
    com o formulário de pagamento fixo (antes era um painel retrátil que fechava ao clicar)."""
    estado = _estado(demo)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/")
        await user.should_see("Competência 02/2023", retries=ESPERA)
        user.find(marker="ver-darf").click()
        await user.should_see("Registrar o pagamento", retries=ESPERA)
    assert estado.ano == 2023 and estado.darf_escolhida is None


@pytest.mark.asyncio
async def test_painel_mostra_darf_pendencias_e_grafico(demo: Path) -> None:
    estado = _estado(demo)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/")
        await user.should_see("Painel de 2026")
        await user.should_see("Competência 02/2023", retries=ESPERA)  # atraso de outro ano
        await user.should_see("DARF(s) em atraso (2023")
        await user.should_not_see("aviso(s) para conferir no ano")  # venda a descoberto: sem aviso
        grafico = user.find(marker="grafico-resultado").elements.pop()
        assert isinstance(grafico, ui.echart)
        assert [s["name"] for s in grafico.options["series"]] == ["Comum", "Day trade", "FII"]


@pytest.mark.asyncio
async def test_cadastros_adiciona_e_renomeia_titular(tmp_path: Path) -> None:
    estado = _estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Primeiro")
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/cadastros")
        await user.should_see("Primeiro")
        await user.should_see("CPF ***.***.247-** · id 1")  # nunca o CPF inteiro
        user.find(marker="novo-titular").click()
        user.find(marker="novo-nome").type("Segundo")
        user.find(marker="novo-cpf").type(cpf_de_teste("111444777"))
        user.find(marker="cadastrar-titular").click()
        await user.should_see("Titular cadastrado.", retries=ESPERA)
        user.find(marker="renomear-1").click()
        user.find(marker="nome-novo").clear().type("Primeiro Renomeado")
        user.find(marker="salvar-nome").click()
        await user.should_see("Primeiro Renomeado")
    assert sorted(t.nome for t in estado.titulares()) == ["Primeiro Renomeado", "Segundo"]


@pytest.mark.asyncio
async def test_configuracoes_backup_restaurar_e_tema(
    tmp_path: Path, configuracao_temporaria: Path
) -> None:
    estado = _estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Antes")
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/configuracoes")
        await user.should_see("Nenhum backup ainda.")
        user.find(marker="fazer-backup").click()
        await user.should_see("Backup feito", retries=ESPERA)
        from coir.servicos.cadastros import renomear_titular

        renomear_titular(estado.banco, 1, "Depois")
        estado.agora = lambda: datetime(2026, 10, 16, 12, 5, 0)  # outro nome de backup
        [copia] = manutencao.backups(estado.banco)
        user.find(marker=f"restaurar-{copia.caminho.name}").click()
        user.find(marker="confirmar-restaurar").click()
        await user.should_see("Banco restaurado", retries=ESPERA)
        radio = user.find(marker="tema").elements.pop()
        assert isinstance(radio, ui.radio)
        radio.set_value("escuro")
        await user.should_see("Configurações")
    assert [t.nome for t in estado.titulares()] == ["Antes"]
    assert len(list((tmp_path / "backups").glob("coir-*.db"))) == 2  # o de antes da restauração
    assert json.loads(configuracao_temporaria.read_text(encoding="utf-8"))["tema"] == "escuro"


@pytest.mark.asyncio
async def test_suporte_prepara_confere_e_grava_o_zip(tmp_path: Path) -> None:
    estado = _estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    nome, texto = notas_ficticias.textos(notas_ficticias.carregar())[0]
    nota = notas_ficticias.pdf(texto, tmp_path / f"{nome}.pdf")
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/suporte")
        await user.should_see("Enviar notas para suporte")
        await user.should_see("Dados para ocultar")
        [envio] = user.find(marker="suporte-pdfs").elements
        assert isinstance(envio, ui.upload)
        await envio.handle_uploads(
            [SmallFileUpload(f"{nome}.pdf", "application/pdf", nota.read_bytes())]
        )
        await user.should_see("nota-01: 1 página(s)", retries=ESPERA)
        [zip_] = user.find(marker="suporte-zip").elements
        assert isinstance(zip_, ui.button) and not zip_.enabled  # só depois de conferir
        user.find(marker="suporte-conferi").click()
        assert zip_.enabled
        user.find(marker="suporte-zip").click()
        await user.should_see("Gravado:")
        await user.should_see("suporte.coir@gmail.com")
    assert len(list((estado.dados / "suporte").glob("*/notas-para-suporte-*.zip"))) == 1


@pytest.mark.asyncio
async def test_suporte_grava_os_dados_para_ocultar(tmp_path: Path) -> None:
    from coir.servicos import suporte

    estado = _estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/suporte")
        user.find(marker="ocultar-nome").type("Fulana de Tal")
        user.find(marker="ocultar-cpf").type(cpf_de_teste())
        user.find(marker="ocultar-cidade").type("Cidade Exemplo")
        user.find(marker="ocultar-gravar").click()
        await user.should_see("Dados para ocultar gravados.")
    [d] = suporte.carregar_dados(estado.dados)
    assert (d.nome, d.cpf, d.cidade) == ("Fulana de Tal", cpf_de_teste(), "Cidade Exemplo")


@pytest.mark.asyncio
async def test_importar_troca_a_pasta_das_notas(
    tmp_path: Path, configuracao_temporaria: Path
) -> None:
    """A pasta das notas pode ser outra (a que o usuário já usa); fica na configuração."""
    estado = _estado(tmp_path / "dados" / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    minhas = tmp_path / "minhas-notas"
    minhas.mkdir()
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/importar")
        user.find(marker="pasta_notas-campo").clear().type(str(minhas))
        user.find(marker="pasta_notas-usar").click()
        await user.should_see("Pasta trocada.")
    assert estado.pasta_notas == minhas.resolve()
    assert json.loads(configuracao_temporaria.read_text(encoding="utf-8"))["pasta_notas"]


@pytest.mark.asyncio
async def test_cadastros_grava_o_prejuizo_inicial(tmp_path: Path) -> None:
    """Quem começa sem as notas antigas informa o prejuízo a compensar em 31/12."""
    from decimal import Decimal

    from coir.servicos import saldos_iniciais

    estado = _estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/cadastros")
        user.find(marker="aba-saldos").click()
        user.find(marker="saldo-comum").type("1.234,56")
        user.find(marker="saldo-fii").type("10")
        user.find(marker="salvar-saldos").click()
        await user.should_see("Saldos iniciais gravados.", retries=ESPERA)
    saldo = saldos_iniciais.ler(estado.banco, 1)
    assert saldo is not None
    assert (saldo.ano, saldo.comum, saldo.day_trade, saldo.fii) == (
        2025, Decimal("1234.56"), Decimal("0"), Decimal("10"))  # fmt: skip
