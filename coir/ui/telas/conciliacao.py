"""Tela 15 — Conciliação com a posição da B3: em cada data com relatório consolidado importado,
a posição do COIR (na data de liquidação) contra a da B3, com a causa provável das diferenças."""

from collections.abc import Callable

from nicegui import run, ui

from coir.erros import CoirError
from coir.reports.modelo import Coluna, TipoColuna, Valor
from coir.servicos.conciliacao import Conciliacao
from coir.servicos.estado import Estado
from coir.ui import tabela

T = TipoColuna
COLUNAS = (
    Coluna("Corretora"),
    Coluna("Ativo"),
    Coluna("B3", T.QUANTIDADE),
    Coluna("COIR", T.QUANTIDADE),
    Coluna("Diferença", T.QUANTIDADE),
    Coluna("Causa provável"),
)


def _valores(c: Conciliacao, so_divergencias: bool) -> list[tuple[Valor, ...]]:
    linhas = c.divergentes if so_divergencias else c.linhas
    return [
        (x.corretora, x.codigo, x.b3, x.coir, x.diferenca, x.causa or "confere") for x in linhas
    ]


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Conciliação com a B3").classes("coir-titulo-pagina")
    ui.label(
        "A posição calculada pelo COIR contra a posição da B3 (relatórios consolidados importados "
        "em Importar > Planilhas da B3), por corretora e ativo. A B3 mostra a custódia já "
        "liquidada: as ações negociadas nos últimos pregões antes da data entram depois (D+2); as "
        "ações emprestadas (BTC) continuam suas."
    ).classes("coir-subtitulo")
    corpo = ui.column().classes("w-full gap-3")

    async def carregar() -> None:
        with corpo:
            espera = ui.spinner(size="lg")
        try:
            lista = await run.io_bound(estado.conciliacoes)
        except CoirError as erro:
            corpo.clear()
            with corpo:
                ui.label(f"Não foi possível conciliar: {erro}").classes("coir-negativo")
            return
        espera.delete()
        if lista is None:
            return
        with corpo:
            if not lista:
                ui.label(
                    "Nenhuma posição da B3 importada. Exporte os relatórios consolidados "
                    "(anual de cada ano) na Área do Investidor e importe em Importar.",
                ).classes("coir-vazio w-full")
                return
            _conteudo(lista)

    ui.timer(0.01, carregar, once=True)


def _conteudo(lista: list[Conciliacao]) -> None:
    escolha = {"indice": len(lista) - 1, "so": True}
    with ui.row().classes("w-full items-center gap-2").mark("resumo-conciliacao"):
        botoes = []
        for i, c in enumerate(lista):
            n = len(c.divergentes)
            rotulo = f"{c.data:%d/%m/%Y} {'✓' if not n else f'⚠ {n}'}"
            botao = (
                ui.button(rotulo, on_click=_escolher(i, lambda i: mostrar(i)))
                .props(
                    f"{'outline' if n else 'flat'} dense no-caps"
                    f" color={'warning' if n else 'positive'}"
                )
                .mark(f"data-{c.data.isoformat()}")
            )
            botoes.append(botao)
        ui.space()
        ui.checkbox(
            "Só as diferenças", value=True, on_change=lambda e: alternar(bool(e.value))
        ).mark("so-diferencas")
    detalhe = ui.column().classes("w-full gap-2")

    def mostrar(i: int) -> None:
        escolha["indice"] = i
        c = lista[i]
        detalhe.clear()
        with detalhe:
            ok = len(c.linhas) - len(c.divergentes)
            ui.label(f"Posição em {c.data:%d/%m/%Y}: {ok} de {len(c.linhas)} conferem").classes(
                "coir-titulo text-base"
            )
            valores = _valores(c, bool(escolha["so"]))
            if valores:
                tabela.grade(COLUNAS, valores, altura="h-[480px]", pelo_conteudo=True).mark(
                    "grade-conciliacao"
                )
            else:
                ui.label("Tudo confere nesta data.").classes("coir-positivo")
            for aviso in c.avisos:
                ui.label(aviso).classes("coir-carimbo text-sm")
            if c.divergentes:
                ui.label(
                    "Ajuste de série de opção (desdobramento, grupamento ou bonificação do "
                    "ativo-objeto) e troca de código da série entram em Eventos."
                ).classes("coir-subtitulo text-xs")

    def alternar(so: bool) -> None:
        escolha["so"] = so
        mostrar(int(escolha["indice"]))

    mostrar(len(lista) - 1)


def _escolher(i: int, f: Callable[[int], None]) -> Callable[[], None]:
    return lambda: f(i)
