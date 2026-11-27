"""Olho do topo: esconde os valores em R$ da tela (só apresentação), lembrado entre aberturas.

Quantidade, preço médio, preço e percentual continuam à mostra; campos de digitação e o que se
copia para o programa IRPF ficam com o valor real (escolhas do usuário em 27/11/2026)."""

import json
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest
from nicegui import ui
from nicegui.testing.user_interaction import UserInteraction
from nicegui.testing.user_simulation import user_simulation

from coir.reports.modelo import TipoColuna
from coir.servicos import darf
from coir.servicos.cadastros import adicionar_titular
from coir.ui import mascara
from coir.ui.app import montar
from tests.apoio import cpf_de_teste
from tests.unit.test_servicos_darf import MESES
from tests.unit.ui.test_telas import AGORA, ESPERA, _estado
from tools.banco_demo import criar

M = mascara.MASCARA


@pytest.fixture(scope="module")
def demo(tmp_path_factory: pytest.TempPathFactory) -> Path:
    return criar(tmp_path_factory.mktemp("demo-mascara"))


def _ocultar(configuracao: Path) -> None:
    configuracao.write_text(json.dumps({mascara.CHAVE: "sim"}), encoding="utf-8")


def _gravado(configuracao: Path) -> str | None:
    return json.loads(configuracao.read_text(encoding="utf-8")).get(mascara.CHAVE)


@pytest.mark.asyncio
async def test_mascara_do_dinheiro_dos_textos_das_grades_e_dos_graficos(
    configuracao_temporaria: Path,
) -> None:
    """As funções conferidas dentro de uma janela (a escolha vale por janela)."""
    falhas: list[AssertionError] = []

    def pagina() -> None:
        try:
            _conferir(configuracao_temporaria)
        except AssertionError as erro:
            falhas.append(erro)
        ui.label("conferido")

    async with user_simulation(pagina) as user:
        await user.open("/")
        await user.should_see("conferido")
    if falhas:
        raise falhas[0]


def _conferir(configuracao_temporaria: Path) -> None:
    frase = "excedente de R$ 1.234,56 (R$ -0,50 de ajuste), limite de 20 mil"
    assert not mascara.ocultos()
    assert mascara.texto(frase) == frase
    assert mascara.dinheiro(Decimal("1234.5")) == "1.234,50"
    assert mascara.alternar() and _gravado(configuracao_temporaria) == "sim"
    assert mascara.texto(frase) == f"excedente de R$ {M} (R$ {M} de ajuste), limite de 20 mil"
    assert mascara.texto("R$ 1234,56 e R$ 12,3456; isento até R$ 35 mil. Pago R$ 100.") == (
        f"R$ {M} e R$ {M}; isento até R$ {M} mil. Pago R$ {M}."
    )
    assert mascara.dinheiro(Decimal("1234.5")) == M
    # nas grades só o dinheiro leva a máscara; quantidade, PM e percentual ficam
    assert mascara.celula(Decimal("-10.5"), TipoColuna.DINHEIRO) == M
    assert mascara.celula(Decimal(100), TipoColuna.QUANTIDADE) == "100"
    assert mascara.celula(Decimal("12.3456789"), TipoColuna.PM) == "12,34567890"
    assert mascara.celula(Decimal("0.15"), TipoColuna.PERCENTUAL) == "15,00%"
    assert mascara.celula(None, TipoColuna.DINHEIRO) == ""
    opcoes = {"tooltip": {"trigger": "axis"}, "xAxis": {"type": "category"},
              "yAxis": {"type": "value", "axisLabel": {"color": "#000"}}}  # fmt: skip
    oculto = mascara.grafico(opcoes)
    assert oculto["yAxis"]["axisLabel"] == {"color": "#000", ":formatter": f"() => '{M}'"}
    assert oculto["tooltip"][":valueFormatter"] == f"() => '{M}'"
    assert oculto["xAxis"] == {"type": "category"} and "axisLabel" in opcoes["yAxis"]
    assert opcoes["yAxis"]["axisLabel"] == {"color": "#000"}  # o original não muda
    assert not mascara.alternar() and _gravado(configuracao_temporaria) == "nao"
    assert mascara.grafico(opcoes) is opcoes


@pytest.mark.asyncio
async def test_olho_do_topo_esconde_mostra_e_grava(
    demo: Path, configuracao_temporaria: Path
) -> None:
    estado = _estado(demo)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/")
        await user.should_see("Competência 02/2023", retries=ESPERA)
        await user.should_not_see(M)
        user.find(marker="olho").click()
        await user.should_see(M, retries=ESPERA)  # DARF da vez e números do ano
        await user.should_see("Competência 02/2023", retries=ESPERA)  # o texto continua
        grafico = user.find(marker="grafico-resultado").elements.pop()
        assert isinstance(grafico, ui.echart)
        assert grafico.options["yAxis"]["axisLabel"][":formatter"] == f"() => '{M}'"
        assert _gravado(configuracao_temporaria) == "sim"
        user.find(marker="olho").click()
        await user.should_see("Competência 02/2023", retries=ESPERA)
        await user.should_not_see(M)
        assert _gravado(configuracao_temporaria) == "nao"


@pytest.mark.asyncio
async def test_lembrado_ao_abrir_e_grades_so_escondem_dinheiro(
    demo: Path, configuracao_temporaria: Path
) -> None:
    _ocultar(configuracao_temporaria)
    estado = _estado(demo)
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/operacoes")
        grade = user.find(marker="grade-operacoes").elements.pop()
        assert isinstance(grade, ui.aggrid)
        nomes = [c["headerName"] for c in grade.options["columnDefs"]]
        total, custos, qtd = (f"c{nomes.index(n)}" for n in ("Total", "Custos", "Quantidade"))
        linhas = grade.options["rowData"]
        assert linhas and {r[total] for r in linhas} == {M} and {r[custos] for r in linhas} == {M}
        assert all(r[qtd] != M for r in linhas)
        # o número bruto continua escondido na linha: a ordenação por valor segue funcionando
        assert all(Decimal(r[f"n{nomes.index('Total')}"]) for r in linhas)
        assert grade.options["pinnedBottomRowData"][0][total] == f"C {M} · V {M}"


@pytest.mark.asyncio
async def test_darf_oculta_mas_o_campo_do_pagamento_tem_o_valor_real(
    tmp_path: Path, configuracao_temporaria: Path
) -> None:
    _ocultar(configuracao_temporaria)
    estado = _estado(tmp_path / "coir.db")
    estado.agora = lambda: datetime(2026, 4, 2, 10, 0, 0)
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    estado.darfs = lambda: darf.darfs(  # type: ignore[method-assign]
        MESES, darf.pagamentos(estado.banco, 1), date(2026, 4, 2), estado.fiscais
    )
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/darf")
        await user.should_see("2 DARF(s) em atraso", retries=ESPERA)
        grade = user.find(marker="grade-darf").elements.pop()
        assert isinstance(grade, ui.aggrid)
        fevereiro = next(r for r in grade.options["rowData"] if r["c0"] == "02/2026")
        assert M in fevereiro.values()
        user.find(marker="grade-darf").trigger("rowClicked", {"data": fevereiro})
        await user.should_see("DARF de fevereiro/2026")
        await user.should_see(M)
        campo = user.find(marker="valor-pagamento").elements.pop()
        assert isinstance(campo, ui.input) and campo.value == "93,10"
        assert "coir-campo-oculto" in campo.classes  # pontos na tela até o clique (CSS)
        assert user.find("93,10").elements == {campo}  # só no campo, não no documento
        user.find(marker="data-pagamento").clear().type("2026-04-01")
        user.find(marker="salvar-pagamento").click()
        await user.should_see(f"Pago em 01/04/2026: {M}", retries=ESPERA)
    pago = darf.pagamentos(estado.banco, 1)[(2026, 2)]
    # a data nova recalculou o valor sugerido (multa e juros até 01/04): real, não a máscara
    assert pago.valor == Decimal("92.80")


@pytest.mark.asyncio
async def test_declaracao_mostra_a_mascara_e_copia_o_valor_real(
    demo: Path, configuracao_temporaria: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _ocultar(configuracao_temporaria)
    copiados: list[str] = []
    monkeypatch.setattr(ui.clipboard, "write", copiados.append)
    estado = _estado(demo)
    estado.titular, estado.ano = 1, 2025
    async with user_simulation(lambda: montar(estado)) as user:
        await user.open("/declaracao")
        await user.should_see(f"31/12/2025: {M}", retries=ESPERA)
        rotulo = next(e for e in user.find(f"31/12/2025: {M}").elements)
        assert rotulo.parent_slot is not None
        [botao] = [e for e in rotulo.parent_slot.parent if isinstance(e, ui.button)]
        UserInteraction(user, {botao}, None).click()
        await user.should_see(f"Copiado: 31/12/2025: {M}")
    assert copiados and copiados[0] != M and "," in copiados[0]
    assert Decimal(copiados[0].replace(",", ".")) >= 0
