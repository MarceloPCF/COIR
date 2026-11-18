"""Tela Simulação sobre o banco de demonstração (notas fictícias, sem navegador): a posição vira
ordem com um clique, o resultado sai lado a lado e nada é gravado no banco."""

from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
from nicegui import ui
from nicegui.testing.user_simulation import user_simulation

from coir.servicos import simulacao
from coir.servicos.simulacao import OrdemDigitada
from coir.ui.app import montar
from coir.ui.telas.simulacao import _Ordem, digitadas
from tests.unit.ui.test_telas import ESPERA, _estado
from tools.banco_demo import criar

D = Decimal
DIA = date(2026, 10, 16)  # o relógio dos testes da interface; última nota em 14/08/2026


@pytest.fixture(scope="module")
def demo(tmp_path_factory: pytest.TempPathFactory) -> Path:
    return criar(tmp_path_factory.mktemp("demo-simulacao"))


def test_linhas_da_tela_viram_ordens_de_cada_cenario() -> None:
    """Sem data no cenário, a ordem fica fora dele (execução parcial); valores em pt-BR."""
    a = _Ordem(ticker="ggbr4", conta=1, quantidade="100", preco="65,50", custos="",
               datas={"A": "2026-10-16", "B": ""})  # fmt: skip
    b = _Ordem(ticker="GGBR4", conta=1, quantidade="100", preco="66", custos="1,20",
               datas={"A": "2026-10-16", "B": "2026-11-16"})  # fmt: skip
    vazia = _Ordem(datas={"A": "", "B": ""})
    cenarios = digitadas([a, b, vazia])
    assert [o.numero for o in cenarios["A"]] == [1, 2]
    assert [(o.numero, o.data) for o in cenarios["B"]] == [(2, date(2026, 11, 16))]
    assert cenarios["A"][0].preco == D("65.50") and cenarios["B"][0].custos == D("1.20")
    with pytest.raises(ValueError, match="ordem 1: quantidade"):
        digitadas([_Ordem(conta=1, quantidade="x", preco="1", datas={"A": "2026-10-16", "B": ""})])
    with pytest.raises(ValueError, match="escolha a conta"):
        digitadas([_Ordem(quantidade="1", preco="1", datas={"A": "2026-10-16", "B": ""})])


def test_servico_compara_cenarios_sem_gravar(demo: Path) -> None:
    """300 GGBR4 da conta da Ágora: vender tudo em outubro (A) ou metade em outubro e metade em
    novembro (B). O banco não muda; os dois cenários vão até novembro."""
    estado = _estado(demo)
    antes = demo.read_bytes()
    pm = next(p.pm for p in simulacao.contexto(demo, estado.refs, estado.fiscais, 1, DIA).posicoes)

    def ordem(numero: int, quantidade: int, dia: date) -> OrdemDigitada:
        return OrdemDigitada(numero=numero, ticker="GGBR4", conta=1, cv="V",
                             quantidade=D(quantidade), preco=D(70), custos=D(0),
                             data=dia)  # fmt: skip

    r = simulacao.simular_banco(demo, estado.refs, estado.fiscais, 1, {
        "A": [ordem(1, 150, DIA), ordem(2, 150, DIA)],
        "B": [ordem(1, 150, DIA), ordem(2, 150, date(2026, 11, 16))],
    })  # fmt: skip
    assert demo.read_bytes() == antes
    assert r["A"].ate == r["B"].ate == date(2026, 11, 30)
    ganho = (D(70) - pm) * 300
    assert sum((o.resultado for o in r["A"].ordens), D(0)) == pytest.approx(ganho, abs=D("0.01"))
    # 21.000,00 de vendas: o limite pesa só para o cenário A (10.500,00 por mês no B)
    assert r["A"].isento() == 0 and r["B"].isento() > 0
    assert r["A"].posicoes[0].quantidade == r["B"].posicoes[0].quantidade == 0


def _grade(user: object, marca: str) -> ui.aggrid:
    grade = user.find(marker=marca).elements.pop()  # type: ignore[attr-defined]
    assert isinstance(grade, ui.aggrid)
    return grade


@pytest.mark.asyncio
async def test_clique_na_posicao_cria_a_venda_e_simular_mostra_o_quadro(demo: Path) -> None:
    estado = _estado(demo)
    estado.titular, estado.ano = 1, 2026
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/simulacao")
        await user.should_see("Simulação informativa", retries=ESPERA)
        await user.should_see("Notas importadas até 14/08/2026", retries=ESPERA)
        posicoes = _grade(user, "sim-posicoes")
        linha = posicoes.options["rowData"][0]
        assert linha["c0"] == "GGBR4"
        user.find(marker="sim-posicoes").trigger("rowClicked", {"data": linha})
        await user.should_see(marker="sim-preco-1")
        user.find(marker="sim-preco-1").type("70")
        user.find(marker="sim-data-b-1").clear()  # fora do cenário B: só o A é simulado
        user.find(marker="sim-simular").click()
        await user.should_see("Resultado até 10/2026", retries=ESPERA)
        await user.should_see(marker="sim-comparacao")
        await user.should_see("Imposto no período")
        await user.should_see("GGBR4: quantidade e custo que ficam")
        ordens = _grade(user, "sim-ordens-A").options["rowData"]
        assert [o["c1"] for o in ordens] == ["GGBR4"]
        await user.should_not_see(marker="sim-ordens-B")


@pytest.mark.asyncio
async def test_venda_acima_da_posicao_e_recusada(demo: Path) -> None:
    estado = _estado(demo)
    estado.titular, estado.ano = 1, 2026
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/simulacao")
        await user.should_see("Notas importadas até", retries=ESPERA)
        user.find(marker="sim-adicionar").click()
        user.find(marker="sim-ticker-1").type("GGBR4")
        user.find(marker="sim-conta-1").elements.pop().set_value(1)  # type: ignore[attr-defined]
        user.find(marker="sim-quantidade-1").type("301")
        user.find(marker="sim-preco-1").type("70")
        user.find(marker="sim-simular").click()
        await user.should_see("passa da posição da conta", retries=ESPERA)
        await user.should_not_see("Resultado até")
