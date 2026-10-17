"""F6: a janela abre com barra lateral, titular e ano; toda tela da navegação monta sem erro."""

from datetime import date, datetime
from pathlib import Path

import pytest
from nicegui.testing.user_simulation import user_simulation

from coir.servicos.estado import Estado
from coir.ui.app import TELAS, montar
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
        await user.should_see("As DARFs do ano, o vencimento, o pagamento e a multa por atraso.")
