"""Tela 4 — Notas: a lista do ano com filtros e, ao escolher uma, a nota como foi lida (negócios,
custos e conciliação) ao lado da página do PDF original."""

import base64
from collections.abc import Callable
from decimal import Decimal

from nicegui import run, ui

from coir.erros import CoirError
from coir.reports.fiscais import MESES
from coir.reports.modelo import Coluna, TipoColuna, Valor, formatar
from coir.servicos import notas
from coir.servicos.estado import Estado
from coir.ui import tabela

T = TipoColuna
COLUNAS_LISTA = (
    Coluna("Data", T.DATA),
    Coluna("Corretora"),
    Coluna("Número"),
    Coluna("Líquido", T.DINHEIRO),
)
COLUNAS_OPS = (
    Coluna("C/V"),
    Coluna("Mercado"),
    Coluna("Ativo"),
    Coluna("Qtd.", T.QUANTIDADE),
    Coluna("Preço", T.PRECO),
    Coluna("Total", T.DINHEIRO),
    Coluna("DT"),
)


def _dinheiro(valor: Decimal) -> str:
    return formatar(valor, T.DINHEIRO)


def _filtrar(
    lista: list[notas.ResumoNota], mes: int, corretora: str, conta: str, numero: str
) -> list[notas.ResumoNota]:
    return [
        n
        for n in lista
        if (not mes or n.data.month == mes)
        and (not corretora or n.corretora == corretora)
        and (not conta or n.conta == conta)
        and numero.strip() in n.numero
    ]


def _contagem(n: int) -> str:
    return f"{n} nota" if n == 1 else f"{n} notas"


def montar(estado: Estado, ao_mudar: Callable[[], object], nota: int | None = None) -> None:
    ui.label("Notas").classes("coir-titulo-pagina")
    ui.label(
        f"As notas de {estado.ano} gravadas no banco. Escolha uma para conferir com o PDF.",
    ).classes("coir-subtitulo")
    todas = estado.lista_notas()
    if not todas:
        ui.label("Nenhuma nota neste ano. Importe as notas em Importar notas.").classes(
            "coir-vazio w-full"
        )
        return

    filtro = {"mes": 0, "corretora": "", "conta": "", "numero": ""}
    with ui.row().classes("w-full items-center gap-3"):
        ui.select(
            {0: "Todos os meses", **{i + 1: m for i, m in enumerate(MESES)}},
            value=0,
            label="Mês",
            on_change=lambda e: mudar("mes", e.value),
        ).props("outlined dense options-dense").classes("w-[170px]")
        corretoras = sorted({n.corretora for n in todas})
        ui.select(
            {"": "Todas", **{c: c for c in corretoras}},
            value="",
            label="Corretora",
            on_change=lambda e: mudar("corretora", e.value),
        ).props("outlined dense options-dense").classes("w-[150px]")
        contas = sorted({n.conta for n in todas})
        if len(contas) > 1:
            ui.select(
                {"": "Todas", **{c: c for c in contas}},
                value="",
                label="Conta",
                on_change=lambda e: mudar("conta", e.value),
            ).props("outlined dense options-dense").classes("w-[170px]")
        ui.input("Número", on_change=lambda e: mudar("numero", e.value or "")).props(
            "outlined dense clearable"
        ).classes("w-[140px]").mark("filtro-numero")
        contador = ui.label("").classes("coir-subtitulo text-sm")

    def valores(lista: list[notas.ResumoNota]) -> list[tuple[Valor, ...]]:
        return [(n.data, n.corretora, n.numero, n.liquido) for n in lista]

    with ui.row().classes("w-full items-start gap-4 no-wrap max-[1100px]:flex-wrap"):
        with ui.column().classes("w-[340px] shrink-0 gap-2"):
            lista = tabela.grade(
                COLUNAS_LISTA,
                valores(todas),
                chaves=[n.id for n in todas],
                larguras=(90, 62, 84, 100),
                altura="h-[640px]",
                ao_clicar=lambda chave: abrir(int(str(chave))),
            )
            lista.mark("lista-notas")
        detalhe = ui.column().classes("grow min-w-0 gap-3")

    def mudar(chave: str, valor: object) -> None:
        filtro[chave] = valor  # type: ignore[assignment]
        vistas = _filtrar(
            todas,
            int(filtro["mes"]),
            str(filtro["corretora"]),
            str(filtro["conta"]),
            str(filtro["numero"]),
        )
        tabela.atualizar(lista, COLUNAS_LISTA, valores(vistas), chaves=[n.id for n in vistas])
        contador.set_text(f"{len(vistas)} de {_contagem(len(todas))}")

    async def abrir(id_nota: int) -> None:
        try:
            d = await run.io_bound(notas.detalhe, estado.banco, estado.dados, id_nota)
        except CoirError as erro:
            ui.notify(str(erro), type="negative")
            return
        if d is None:
            return
        detalhe.clear()
        with detalhe:
            _detalhe(d)

    contador.set_text(_contagem(len(todas)))
    with detalhe:
        ui.label("Escolha uma nota na lista.").classes("coir-vazio w-full")
    if nota is not None:
        ui.timer(0.01, lambda: abrir(nota), once=True)


def _detalhe(d: notas.DetalheNota) -> None:
    r = d.resumo
    with ui.row().classes("w-full items-baseline gap-3"):
        ui.label(f"Nota {r.numero}").classes("coir-titulo text-xl")
        ui.label(
            f"{r.corretora} · conta {r.conta} · {r.data:%d/%m/%Y} · "
            f"{'Bovespa' if r.segmento == 'bovespa' else 'BM&F'}"
        ).classes("coir-subtitulo")
    with ui.row().classes("w-full items-start gap-4"):
        with ui.column().classes("min-w-[300px] flex-1 gap-3"):
            with ui.row().classes("items-start gap-2 no-wrap"):
                if d.conciliada:
                    ui.icon("verified", size="20px").classes("coir-positivo")
                    ui.label("Conciliada: negócios, custos e líquido fecham ao centavo.").classes(
                        "coir-positivo text-sm"
                    )
                else:
                    ui.icon("error", size="20px").classes("coir-negativo")
                    ui.label(f"Não concilia: {d.motivo}").classes("coir-negativo text-sm")
            with ui.grid(columns="1fr auto auto").classes("coir-folha w-full gap-x-4 gap-y-1"):
                ui.label("Resumo").classes("coir-campo-rotulo")
                ui.label("Impresso").classes("coir-campo-rotulo text-right")
                ui.label("Calculado").classes("coir-campo-rotulo text-right")
                for c in d.conferencias:
                    ui.label(c.rotulo)
                    ui.label(_dinheiro(c.impresso)).classes("coir-numero text-right")
                    if c.calculado is None:
                        ui.label("")
                    else:
                        classe = "coir-positivo" if c.confere else "coir-negativo"
                        marca = "✓" if c.confere else "✗"
                        ui.label(f"{_dinheiro(c.calculado)} {marca}").classes(
                            f"coir-numero text-right {classe}"
                        )
            resumo = d.nota.resumo_bovespa or d.nota.resumo_bmf
            assert resumo is not None
            custos = [c for c in resumo.custos if c.valor]
            ui.label("Custos (itens zerados omitidos)").classes("coir-campo-rotulo")
            with ui.grid(columns="1fr auto").classes("coir-folha w-full gap-x-4 gap-y-0 py-2"):
                for c in custos:
                    ui.label(c.rotulo).classes("text-sm")
                    ui.label(_dinheiro(c.valor)).classes("coir-numero text-right text-sm")
        with ui.column().classes("min-w-[380px] flex-[1.1] gap-2"):
            _pdf(d)
    ops: list[tuple[Valor, ...]] = [
        (
            o.cv,
            o.mercado,
            o.papel,
            o.quantidade,
            o.preco,
            o.total,
            "DT" if o.categoria == "DAY_TRADE" else "",
        )
        for o in d.operacoes
    ]
    ui.label(f"Operações ({len(ops)})").classes("coir-campo-rotulo")
    tabela.grade(COLUNAS_OPS, ops, altura="h-[300px]")


def _pdf(d: notas.DetalheNota) -> None:
    if d.pdf is None:
        ui.label(
            f"O PDF {d.resumo.arquivo} não está na pasta das notas: a nota fica no banco, "
            "mas sem o original para conferir."
        ).classes("coir-vazio w-full")
        return
    pdf = d.pdf
    paginas: tuple[int, ...] = d.resumo.paginas or (1,)
    with ui.row().classes("items-center gap-2"):
        ui.label(f"{d.resumo.arquivo}").classes("coir-subtitulo text-sm break-all")
        botoes = ui.row().classes("gap-1")
    imagem = ui.image().classes("coir-folha w-full p-0 cursor-zoom-in").mark("pagina-pdf")
    with ui.dialog().props("maximized") as ampliar, ui.card().classes("w-full items-center"):
        grande = ui.image().classes("max-w-[1400px] w-full")
        ui.button("Fechar", on_click=ampliar.close).props("flat")
    imagem.on("click", ampliar.open)

    async def mostrar(pagina: int) -> None:
        try:
            png = await run.io_bound(notas.pagina_png, pdf, pagina)
        except CoirError as erro:
            ui.notify(str(erro), type="negative")
            return
        if png is None:
            return
        fonte = "data:image/png;base64," + base64.b64encode(png).decode()
        imagem.set_source(fonte)
        grande.set_source(fonte)

    def botao(pagina: int) -> None:
        ui.button(f"pág. {pagina}", on_click=lambda: mostrar(pagina)).props("flat dense size=sm")

    primeira = paginas[0]
    with botoes:
        for p in paginas if len(paginas) > 1 else ():
            botao(p)
    ui.timer(0.01, lambda: mostrar(primeira), once=True)
