"""F6: o estado do app calcula os dados de todas as telas com o banco real (2008-2026) e cada tela
abre sem erro; a apuração inicial leva segundos e as telas seguintes vêm do cache.

Roda numa CÓPIA do coir.db do usuário, com as planilhas da B3 importadas (conftest): nunca migra
nem grava o banco real.
"""

import time
from datetime import datetime
from pathlib import Path

import pytest
from nicegui.testing.user_simulation import user_simulation

from coir.config import VAR_CONFIG
from coir.domain.fiscal import Visao
from coir.servicos.estado import Estado
from coir.ui.app import INICIO, TELAS, montar

pytestmark = [pytest.mark.golden, pytest.mark.slow]
AGORA = datetime(2026, 10, 17, 12, 0, 0)


def _estado(banco: Path) -> Estado:
    estado = Estado(banco)
    estado.agora = lambda: AGORA
    return estado


def test_dados_de_todas_as_telas_em_todos_os_anos(copia: Path) -> None:
    estado = _estado(copia)
    inicio = time.perf_counter()
    estado.apuracao(AGORA.date())
    apuracao_inicial = time.perf_counter() - inicio
    assert apuracao_inicial < 60, f"apuração inicial levou {apuracao_inicial:.1f} s"
    for titular in estado.titulares():
        estado.titular = titular.id
        for ano in estado.anos(AGORA.date()):
            estado.ano = ano
            fim = estado.fim_do_ano()
            estado.apuracao(fim)
            estado.lista_notas()
            estado.livro()
            estado.darfs()
            for visao in Visao:
                estado.carteira(fim, visao)
            estado.movimentos(fim, Visao.TITULAR)
            estado.opcoes()
            estado.futuros()
            if any(m.titular == titular.id and m.ano == ano for m in estado.apuracao(fim).meses):
                estado.dados_relatorio()
        estado.proventos()
        estado.conciliacoes()
    inicio = time.perf_counter()
    estado.livro()
    estado.darfs()
    assert time.perf_counter() - inicio < 0.5  # do cache


# o que cada tela mostra depois de carregar (o rótulo do menu não serve: está sempre na tela)
CONTEUDO = {
    "/": "DARF da vez", "/importar": "Pasta das notas:", "/notas": "Escolha uma nota na lista.",
    "/operacoes": "Clique numa operação", "/apuracao": "Apurado até", "/darf": "DARF",
    "/relatorios": "Resumo fiscal do ano", "/carteira": "posições em", "/cadastros": "CPF ***",
    "/configuracoes": "Registro de atividades", "/inicio": "Bem-vindo ao COIR",
    "/opcoes": "Prêmio das abertas", "/futuros": "Resultado de 2026",
    "/proventos": "Dividendos · isentos (09)", "/outros-custos": "Total do ano:",
    "/conciliacao": "Posição em ",
}  # fmt: skip


@pytest.mark.asyncio
async def test_cada_tela_abre_com_os_dados_reais(
    copia: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv(VAR_CONFIG, str(tmp_path / "config.json"))
    estado = _estado(copia)
    assert set(CONTEUDO) == {t.rota for t in (*TELAS, INICIO)}
    async with user_simulation(lambda: montar(estado)) as user:
        for rota, texto in CONTEUDO.items():
            await user.open(rota)
            await user.should_see(texto, retries=1200)
            await user.should_not_see("produced an error")
            await user.should_not_see("Não foi possível")
