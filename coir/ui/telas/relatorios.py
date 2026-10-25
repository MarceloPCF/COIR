"""Tela 17 — Relatórios: os relatórios fiscais do ano na tela (cada seção numa grade) e a
exportação em xlsx, csv e pdf para a pasta de dados (o nome dos arquivos não leva CPF)."""

import os
import sys
from collections.abc import Awaitable, Callable
from datetime import date
from decimal import Decimal
from pathlib import Path

from nicegui import run, ui

from coir.domain.relatorio import DadosRelatorio
from coir.erros import CoirError
from coir.reports.fiscais import ABAS, MESES, TODOS, relatorios
from coir.reports.modelo import Coluna, Relatorio, Secao, TipoColuna, Valor
from coir.servicos.estado import Estado
from coir.servicos.relatorios import FORMATOS, exportar
from coir.ui import tabela

T = TipoColuna

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
    elif rel.chave == "operacoes-finalizadas":
        _finalizadas(secoes)
    elif len(secoes) > MAX_SECOES:
        _uma_por_vez(rel, secoes)
    else:
        for s in secoes:
            tabela.secao(s, altura="h-[480px]")
    for nota in rel.notas:
        ui.label(nota).classes("coir-subtitulo text-xs")


def _uma_por_vez(rel: Relatorio, secoes: list[Secao]) -> None:
    """Relatório com muitas seções (a conferência do PM tem uma por ativo; a renda variável, uma
    por mês): uma grade por vez, escolhida na lista, ou todas, para imprimir. Dezenas de grades de
    uma vez deixam a página lenta."""
    rotulo = "Ativo" if rel.chave == "conferencia-pm" else "Seção"
    nomes = {i: s.titulo or f"{rotulo} {i + 1}" for i, s in enumerate(secoes)}
    with ui.row().classes("items-center gap-4"):
        escolha = ui.select(nomes, value=0, label=f"{rotulo} ({len(secoes)})", with_input=True)
        escolha.props("outlined dense options-dense").classes("min-w-[260px]")
        escolha.mark("secao-relatorio")
        todas = ui.checkbox("Mostrar todas (para imprimir)").mark("todas-as-secoes")
    area = ui.column().classes("w-full gap-2")

    def ver() -> None:
        area.clear()
        escolha.set_enabled(not todas.value)
        with area:
            for s in secoes if todas.value else [secoes[int(escolha.value or 0)]]:
                tabela.secao(s, altura="h-[520px]")

    escolha.on_value_change(lambda _: ver())
    todas.on_value_change(lambda _: ver())
    ver()


def _finalizadas(secoes: list[Secao]) -> None:
    """Operações finalizadas do ano numa grade só, com filtros (meses, corretora, ativo,
    categoria e tipo) e o resultado somado por categoria e tipo, como nos controles de IR."""
    colunas = secoes[0].colunas
    nome = [c.nome for c in colunas]
    i_data, i_ativo, i_corretora = nome.index("Data"), nome.index("Ativo"), nome.index("Corretora")
    i_cat, i_tipo, i_res = nome.index("Categoria"), nome.index("Tipo"), nome.index("Resultado")
    todas = [linha for s in secoes for linha in s.linhas]
    filtro: dict[str, object] = {"de": 1, "ate": 12, "corretora": "", "ativo": "", "cat": "",
                                 "tipo": ""}  # fmt: skip

    def vistas() -> list[tuple[Valor, ...]]:
        ativo = str(filtro["ativo"]).strip().upper()
        de, ate = int(str(filtro["de"])), int(str(filtro["ate"]))
        saida: list[tuple[Valor, ...]] = []
        for x in todas:
            dia = x[i_data]
            if not (isinstance(dia, date) and de <= dia.month <= ate):
                continue
            if ativo not in str(x[i_ativo]) or any(
                filtro[k] and x[i] != filtro[k]
                for k, i in (("corretora", i_corretora), ("cat", i_cat), ("tipo", i_tipo))
            ):
                continue
            saida.append(tuple(x))
        return saida

    def soma(linhas: list[tuple[Valor, ...]]) -> Decimal:
        return sum((v for x in linhas if isinstance(v := x[i_res], Decimal)), Decimal(0))

    def total(linhas: list[tuple[Valor, ...]]) -> tuple[Valor, ...]:
        rodape: list[Valor] = [None] * len(colunas)
        rodape[0], rodape[i_res] = f"{len(linhas)} operações", soma(linhas)
        return tuple(rodape)

    def resumo(linhas: list[tuple[Valor, ...]]) -> list[tuple[Valor, ...]]:
        grupos: dict[tuple[str, str], list[tuple[Valor, ...]]] = {}
        for x in linhas:
            grupos.setdefault((str(x[i_cat]), str(x[i_tipo])), []).append(x)
        return [(c, t, len(g), soma(g)) for (c, t), g in sorted(grupos.items())]

    meses = {i + 1: m for i, m in enumerate(MESES)}
    with ui.row().classes("w-full items-center gap-3"):
        for chave, rotulo, valor in (("de", "De", 1), ("ate", "Até", 12)):
            ui.select(meses, value=valor, label=rotulo, on_change=lambda e, k=chave: mudar(
                k, e.value)).props("outlined dense options-dense").classes("w-[140px]")  # fmt: skip
        ui.input("Ativo", on_change=lambda e: mudar("ativo", e.value or "")).props(
            "outlined dense clearable").classes("w-[130px]").mark("finalizadas-ativo")  # fmt: skip
        selecoes = (("corretora", "Corretora", i_corretora), ("cat", "Categoria", i_cat),
                    ("tipo", "Tipo", i_tipo))  # fmt: skip
        for chave, rotulo, i in selecoes:
            opcoes = sorted({str(x[i]) for x in todas})
            ui.select({"": "Todas", **{o: o for o in opcoes}}, value="", label=rotulo,
                      on_change=lambda e, k=chave: mudar(k, e.value or ""),
                      ).props("outlined dense options-dense").classes("w-[160px]").mark(
                f"finalizadas-{chave}")  # fmt: skip
    ui.label("Resultado por categoria e tipo").classes("coir-titulo text-base")
    colunas_resumo = (Coluna("Categoria"), Coluna("Tipo"), Coluna("Operações", T.INTEIRO),
                      Coluna("Resultado", T.DINHEIRO))  # fmt: skip
    linhas = vistas()
    grade_resumo = tabela.grade(colunas_resumo, resumo(linhas), total=("Total", "", len(linhas),
                                soma(linhas)))  # fmt: skip
    grade_resumo.mark("finalizadas-resumo")
    grade = tabela.grade(colunas, linhas, total=total(linhas), altura="h-[560px]",
                         pelo_conteudo=True)  # fmt: skip
    grade.mark("finalizadas")

    def mudar(chave: str, valor: object) -> None:
        filtro[chave] = valor
        linhas = vistas()
        tabela.atualizar(grade, colunas, linhas, total=total(linhas))
        tabela.atualizar(grade_resumo, colunas_resumo, resumo(linhas),
                         total=("Total", "", len(linhas), soma(linhas)))  # fmt: skip
