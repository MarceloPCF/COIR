"""Tela 17 — Relatórios: os relatórios fiscais do ano na tela (cada seção numa grade) e a
exportação em xlsx, csv e pdf para a pasta de dados (o nome dos arquivos não leva CPF)."""

import os
import sys
from collections.abc import Awaitable, Callable
from pathlib import Path

from nicegui import run, ui

from coir.domain.relatorio import DadosRelatorio
from coir.erros import CoirError
from coir.reports.fiscais import ABAS, TODOS, relatorios
from coir.reports.modelo import Relatorio, Secao
from coir.servicos.estado import Estado
from coir.servicos.relatorios import FORMATOS, exportar
from coir.ui import tabela

MAX_SECOES = 8  # acima disso, uma seção por vez


def _abrir_pasta(pasta: Path) -> None:
    if sys.platform == "win32" and pasta.is_dir():
        os.startfile(pasta)
    else:
        ui.notify(str(pasta))


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Relatórios").classes("coir-titulo-pagina")
    ui.label(
        f"Os relatórios de {estado.ano}: escolha um para ver na tela, ou exporte todos em xlsx, "
        "csv e pdf (sem precisar do Excel)."
    ).classes("coir-subtitulo")
    destino_padrao = estado.dados / "relatorios" / str(estado.ano)
    dados: list[DadosRelatorio] = []

    with ui.row().classes("coir-folha w-full items-center gap-4"):
        ui.label("Exportar todos").classes("coir-titulo text-base")
        formatos = {f: ui.checkbox(f, value=True).mark(f"formato-{f}") for f in FORMATOS}
        pasta = (
            ui.input("Pasta", value=str(destino_padrao))
            .props("outlined dense")
            .classes("grow min-w-[320px]")
            .mark("pasta-relatorios")
        )
        botao = ui.button("Exportar", icon="download").mark("exportar")
        ui.button(
            "Abrir a pasta",
            icon="folder_open",
            on_click=lambda: _abrir_pasta(Path(pasta.value or ".")),
        ).props("flat")

    def ao_escolher(chave: str) -> Callable[[], Awaitable[None]]:
        return lambda: mostrar(chave)

    with ui.row().classes("w-full items-start gap-4 no-wrap"):
        with ui.column().classes("w-[240px] shrink-0 gap-0 coir-folha p-1"):
            itens = {
                chave: ui.item(ABAS[chave], on_click=ao_escolher(chave))
                .props("clickable dense")
                .classes("rounded")
                .mark(f"relatorio-{chave}")
                for chave in TODOS
            }
        corpo = ui.column().classes("grow min-w-0 gap-3")

    async def preparar() -> DadosRelatorio | None:
        if dados:
            return dados[0]
        try:
            d = await run.io_bound(estado.dados_relatorio)
        except CoirError as erro:
            ui.notify(f"Não foi possível gerar: {erro}", type="negative", multi_line=True)
            return None
        if d is not None:
            dados[:] = [d]
        return d

    async def mostrar(chave: str) -> None:
        for c, item in itens.items():
            item.classes(
                add="coir-item-ativo" if c == chave else "",
                remove="" if c == chave else "coir-item-ativo",
            )
        corpo.clear()
        with corpo:
            espera = ui.spinner(size="lg")
        d = await preparar()
        if espera.is_deleted:  # outro relatório foi escolhido enquanto este carregava
            return
        if d is None:
            corpo.clear()
            return
        gerados = await run.io_bound(relatorios, d, [chave])
        if gerados is None or espera.is_deleted:
            return
        [rel] = gerados
        espera.delete()
        with corpo:
            _relatorio(rel)

    async def gravar() -> None:
        escolhidos = tuple(f for f, caixa in formatos.items() if caixa.value)
        if not escolhidos:
            ui.notify("Escolha ao menos um formato.", type="warning")
            return
        d = await preparar()
        if d is None:
            return
        botao.set_enabled(False)
        try:
            arquivos = await run.io_bound(exportar, d, Path(pasta.value or "."), escolhidos)
        except (CoirError, OSError) as erro:
            ui.notify(f"Exportação não feita: {erro}", type="negative", multi_line=True)
            return
        finally:
            botao.set_enabled(True)
        if arquivos is not None:
            ui.notify(f"{len(arquivos)} arquivo(s) gravado(s) em {pasta.value}", type="positive")

    botao.on_click(gravar)
    ui.timer(0.01, lambda: mostrar("resumo-fiscal"), once=True)


def _relatorio(rel: Relatorio) -> None:
    c = rel.cabecalho
    with ui.row().classes("items-baseline gap-3"):
        ui.label(rel.titulo).classes("coir-titulo text-xl")
        ui.label(f"{c.titular} · {c.periodo}").classes("coir-subtitulo")
    secoes = [s for s in rel.secoes if s.linhas]
    if not secoes:
        ui.label("Nada neste relatório para o ano.").classes("coir-vazio w-full")
    elif len(secoes) > MAX_SECOES:
        _uma_por_vez(rel, secoes)
    else:
        for s in secoes:
            tabela.secao(s, altura="h-[480px]")
    for nota in rel.notas:
        ui.label(nota).classes("coir-subtitulo text-xs")


def _uma_por_vez(rel: Relatorio, secoes: list[Secao]) -> None:
    """Relatório com muitas seções (a conferência do PM tem uma por ativo): uma grade por vez,
    escolhida na lista. Dezenas de grades de uma vez travam a página ("conexão perdida")."""
    rotulo = "Ativo" if rel.chave == "conferencia-pm" else "Seção"
    nomes = {i: s.titulo or f"{rotulo} {i + 1}" for i, s in enumerate(secoes)}
    escolha = ui.select(nomes, value=0, label=f"{rotulo} ({len(secoes)})", with_input=True)
    escolha.props("outlined dense options-dense").classes("min-w-[260px]").mark("secao-relatorio")
    area = ui.column().classes("w-full gap-2")

    def ver(i: int) -> None:
        area.clear()
        with area:
            tabela.secao(secoes[i], altura="h-[520px]")

    escolha.on_value_change(lambda e: ver(e.value) if e.value is not None else None)
    ver(0)
