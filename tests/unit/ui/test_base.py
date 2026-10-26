"""F6: a janela abre com barra lateral, titular e ano; toda tela da navegação monta sem erro."""

from datetime import date, datetime
from pathlib import Path

import pytest
from nicegui.testing.user_simulation import user_simulation

from coir.servicos.estado import Estado
from coir.ui import tema
from coir.ui.app import TELAS, montar, montar_erro
from tools.banco_demo import NOME, criar

HOJE = date(2026, 10, 16)


def _estado(banco: Path) -> Estado:
    estado = Estado(banco)
    estado.agora = lambda: datetime(2026, 10, 16, 12, 0, 0)
    return estado


@pytest.fixture(scope="module")
def banco(tmp_path_factory: pytest.TempPathFactory) -> Path:
    return criar(tmp_path_factory.mktemp("demo"))


@pytest.mark.asyncio
async def test_janela_mostra_titular_ano_e_navegacao(banco: Path) -> None:
    estado = _estado(banco)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/")
        await user.should_see(NOME)
        await user.should_see("2026")
        for tela in TELAS:
            await user.should_see(tela.rotulo)
    assert (estado.titular, estado.ano) == (1, 2026)


@pytest.mark.asyncio
@pytest.mark.parametrize("rota", [t.rota for t in TELAS])
async def test_cada_tela_abre(banco: Path, rota: str) -> None:
    estado = _estado(banco)
    tela = next(t for t in TELAS if t.rota == rota)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open(rota)
        await user.should_see(tela.rotulo)


@pytest.mark.asyncio
async def test_sem_banco_abre_sem_titular(tmp_path: Path) -> None:
    estado = _estado(tmp_path / "coir.db")
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/")
        await user.should_see("Painel")
    assert estado.titular is None and estado.ano == 2026


@pytest.mark.asyncio
async def test_clicar_na_barra_lateral_troca_de_tela(banco: Path) -> None:
    estado = _estado(banco)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/")
        user.find(marker="menu/darf").click()
        await user.should_see("Nenhuma DARF neste ano.", retries=50)


@pytest.mark.asyncio
async def test_topo_com_a_logomarca(tmp_path: Path) -> None:
    assert tema.simbolo_svg().startswith("<svg") and tema.ICONE.is_file()
    estado = _estado(tmp_path / "coir.db")
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/")
        await user.should_see(marker="marca")
        await user.should_see("Controle de Operações e Imposto de Renda")


@pytest.mark.asyncio
async def test_banco_que_nao_abre_mostra_o_motivo_e_as_pastas(tmp_path: Path) -> None:
    estado = _estado(tmp_path / "coir.db")
    mensagem = "o banco foi gravado por uma versão mais nova do COIR (revisão 9999)"
    async with user_simulation(lambda: montar_erro(estado, mensagem)) as user:
        await user.open("/")
        await user.should_see(mensagem)
        await user.should_see(f"Backups do banco: {tmp_path / 'backups'}")
