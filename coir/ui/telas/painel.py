"""Tela 2 — Painel do ano do titular: a DARF da vez, as pendências, o resultado mês a mês por
categoria (gráfico) e os números do ano, com atalhos para importar e exportar."""

from collections.abc import Callable
from decimal import Decimal

from nicegui import run, ui

from coir.domain.fiscal import Apuracao, Categoria, MesApurado
from coir.erros import CoirError
from coir.reports.fiscais import MESES
from coir.reports.modelo import TipoColuna, formatar
from coir.servicos import darf, notas
from coir.servicos.estado import Estado
from coir.ui import tema

ZERO = Decimal(0)
SERIES = (
    (Categoria.COMUM, "Comum", tema.CLARO.caneta),
    (Categoria.DAY_TRADE, "Day trade", tema.CLARO.carimbo),
    (Categoria.FII, "FII", tema.CLARO.positivo),
)


def _r(valor: Decimal) -> str:
    return formatar(valor, TipoColuna.DINHEIRO)


def _resultado(m: MesApurado, c: Categoria) -> Decimal:
    return next((r.resultado for r in m.categorias if r.categoria is c), ZERO)


def _cartao(titulo: str) -> ui.column:
    with ui.column().classes("coir-folha gap-2 min-w-[300px] flex-1") as cartao:
        ui.label(titulo).classes("coir-campo-rotulo")
    return cartao


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    nome = next((t.nome for t in estado.titulares() if t.id == estado.titular), "")
    ui.label(f"Painel de {estado.ano}").classes("coir-titulo-pagina")
    ui.label(nome).classes("coir-subtitulo")
    corpo = ui.column().classes("w-full gap-4")

    async def carregar() -> None:
        with corpo:
            espera = ui.spinner(size="lg")
        try:
            apuracao = await run.io_bound(estado.apuracao, estado.fim_do_ano())
            darfs = await run.io_bound(estado.darfs)
            incompletos = await run.io_bound(notas.arquivos_incompletos, estado.banco)
        except CoirError as erro:
            corpo.clear()
            with corpo:
                ui.label(f"Não foi possível apurar: {erro}").classes("coir-negativo")
            return
        espera.delete()
        if apuracao is None or darfs is None or incompletos is None:
            return
        with corpo:
            _conteudo(estado, apuracao, darfs, incompletos)

    ui.timer(0.01, carregar, once=True)


def _conteudo(
    estado: Estado, apuracao: Apuracao, darfs: list[darf.LinhaDarf], incompletos: list[str]
) -> None:
    meses = [m for m in apuracao.meses if m.titular == estado.titular and m.ano == estado.ano]
    avisos = [a for a in apuracao.avisos if a[:4] == str(estado.ano) or not a[:4].isdigit()]
    atrasadas = [x for x in darfs if x.situacao is darf.Situacao.EM_ATRASO]
    a_pagar = sorted((x for x in darfs if x.situacao is darf.Situacao.A_PAGAR),
                     key=lambda x: x.vencimento or x.ano)  # fmt: skip
    with ui.row().classes("w-full items-stretch gap-4"):
        with _cartao("DARF da vez").mark("cartao-darf"):
            proxima = (atrasadas or a_pagar or [None])[0]
            if proxima is None:
                ui.label("Nada a pagar neste ano.").classes("coir-titulo text-xl")
            else:
                classe = "coir-negativo" if proxima in atrasadas else "coir-carimbo"
                ui.label(_r(proxima.total)).classes(f"coir-numero text-3xl {classe}")
                vence = f"{proxima.vencimento:%d/%m/%Y}" if proxima.vencimento else ""
                ui.label(f"Competência {proxima.mes:02d}/{proxima.ano} · vence {vence} · "
                         f"{proxima.situacao.value}").classes("text-sm")  # fmt: skip
            ui.button("Ver as DARFs", icon="request_quote",
                      on_click=lambda: ui.navigate.to("/darf")).props("flat dense")  # fmt: skip
        with _cartao("Pendências").mark("cartao-pendencias"):
            itens = [
                (len(atrasadas), "DARF(s) em atraso", "/darf"),
                (len(incompletos), "arquivo(s) com nota rejeitada", "/importar"),
                (len(avisos), "aviso(s) do motor no ano", "/apuracao"),
            ]
            if not any(n for n, _, _ in itens):
                ui.label("Nenhuma pendência.").classes("coir-positivo")
            for n, texto, rota in itens:
                if n:
                    with ui.row().classes("items-center gap-2"):
                        ui.icon("error_outline", size="18px").classes("coir-carimbo")
                        ui.link(f"{n} {texto}", rota).classes("text-sm")
        with _cartao("Atalhos"):
            ui.button("Importar notas", icon="upload_file",
                      on_click=lambda: ui.navigate.to("/importar")).props("outline")  # fmt: skip
            ui.button("Exportar relatórios", icon="description",
                      on_click=lambda: ui.navigate.to("/relatorios")).props("outline")  # fmt: skip
    if not meses:
        ui.label("Nada apurado neste ano: importe as notas.").classes("coir-vazio w-full")
        return
    _numeros(meses, darfs)
    _grafico(meses)


def _numeros(meses: list[MesApurado], darfs: list[darf.LinhaDarf]) -> None:
    ultimo = meses[-1]
    resultado = sum((_resultado(m, c) for m in meses for c, _, _ in SERIES), ZERO)
    imposto = sum((m.imposto_devido for m in meses), ZERO)
    pagas = sum(1 for x in darfs if x.situacao is darf.Situacao.PAGA)
    com_darf = sum(1 for x in darfs if x.situacao is not darf.Situacao.ACUMULADA)
    prejuizo = sum((r.prejuizo_a_compensar for r in ultimo.categorias), ZERO)
    with ui.row().classes("w-full gap-8 coir-folha").mark("numeros-ano"):
        for rotulo, valor in (
            ("Resultado tributável do ano", _r(resultado)),
            ("Imposto devido", _r(imposto)),
            ("DARFs pagas", f"{pagas} de {com_darf}"),
            (f"Prejuízo a compensar ({MESES[ultimo.mes - 1].lower()})", _r(prejuizo)),
        ):
            with ui.column().classes("gap-0"):
                ui.label(valor).classes("coir-numero text-2xl")
                ui.label(rotulo).classes("coir-campo-rotulo")


def _grafico(meses: list[MesApurado]) -> None:
    ui.label("Resultado mês a mês por categoria").classes("coir-titulo text-base")
    tinta = "#7A8792"  # legível no claro e no escuro
    ui.echart({
        "animation": False,
        "grid": {"left": 70, "right": 20, "top": 40, "bottom": 30},
        "legend": {"top": 0, "textStyle": {"color": tinta}},
        "tooltip": {"trigger": "axis", ":valueFormatter":
                    "v => v.toLocaleString('pt-BR', {minimumFractionDigits: 2, "
                    "maximumFractionDigits: 2})"},
        "xAxis": {"type": "category", "data": [MESES[m.mes - 1][:3] for m in meses],
                  "axisLabel": {"color": tinta}},
        "yAxis": {"type": "value", "axisLabel": {"color": tinta, ":formatter":
                  "v => v.toLocaleString('pt-BR')"},
                  "splitLine": {"lineStyle": {"color": "rgba(122,135,146,.25)"}}},
        "series": [
            {"name": nome, "type": "bar", "itemStyle": {"color": cor},
             "data": [float(_resultado(m, c)) for m in meses]}  # só para desenhar
            for c, nome, cor in SERIES
        ],
    }).classes("w-full h-[340px] coir-folha").mark("grafico-resultado")  # fmt: skip
