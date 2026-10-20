"""Tela 2 — Painel do ano do titular: a DARF da vez, as pendências, o resultado mês a mês por
categoria (gráfico) e os números do ano, com atalhos para importar e exportar."""

from collections.abc import Callable
from decimal import Decimal

from nicegui import run, ui

from coir.domain.fiscal import Apuracao, Categoria, MesApurado
from coir.erros import CoirError
from coir.reports.fiscais import MESES
from coir.reports.modelo import TipoColuna, formatar
from coir.servicos import atualizacao, darf, notas, painel
from coir.servicos.estado import Estado
from coir.ui import preferencias, tema

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
    aviso = ui.row().classes("w-full items-center gap-3").mark("aviso-tabelas-novas")
    corpo = ui.column().classes("w-full gap-4")

    async def tabelas_novas() -> None:
        """Opção "verificar ao abrir": uma consulta por execução do app; erro de rede é calado
        aqui (o botão de Configurações mostra o motivo)."""
        if not preferencias.verificar_tabelas():
            return
        if not estado.verificado:
            try:
                estado.remoto = await run.io_bound(atualizacao.verificar, estado.baixar)
            except CoirError:
                return
            finally:
                estado.verificado = True
        if atualizacao.mais_novo(estado.dados, estado.remoto) and estado.remoto is not None:
            with aviso:
                texto = f"Há tabelas de referência novas ({estado.remoto.data:%d/%m/%Y})."
                ui.label(texto).classes("coir-carimbo text-sm")
                ui.button("Ver em Configurações", on_click=lambda: ui.navigate.to(
                    "/configuracoes")).props("flat dense")  # fmt: skip

    async def carregar() -> None:
        with corpo:
            espera = ui.spinner(size="lg")
        try:
            apuracao = await run.io_bound(estado.apuracao, estado.fim_do_ano())
            darfs = await run.io_bound(estado.darfs)
            todas = await run.io_bound(estado.darfs_todos_os_anos)
            incompletos = await run.io_bound(notas.arquivos_incompletos, estado.banco)
        except CoirError as erro:
            corpo.clear()
            with corpo:
                ui.label(f"Não foi possível apurar: {erro}").classes("coir-negativo")
            return
        espera.delete()
        if apuracao is None or darfs is None or todas is None or incompletos is None:
            return
        with corpo:
            _conteudo(estado, apuracao, darfs, todas, incompletos)

    ui.timer(0.01, carregar, once=True)
    ui.timer(0.05, tabelas_novas, once=True)


def _conteudo(
    estado: Estado,
    apuracao: Apuracao,
    darfs: list[darf.LinhaDarf],
    todas: list[darf.LinhaDarf],
    incompletos: list[str],
) -> None:
    """`darfs`: as do ano escolhido (números); `todas`: as de todos os anos (pendências)."""
    meses = [m for m in apuracao.meses if m.titular == estado.titular and m.ano == estado.ano]
    avisos = [a for a in apuracao.avisos if a[:4] == str(estado.ano) or not a[:4].isdigit()]
    atrasadas = [x for x in todas if x.situacao is darf.Situacao.EM_ATRASO]
    with ui.row().classes("w-full items-stretch gap-4"):
        with _cartao("DARF da vez").mark("cartao-darf"):
            vez = painel.da_vez(todas)
            if vez is None:
                ui.label("Nada a pagar.").classes("coir-titulo text-xl")
            else:
                classe = "coir-negativo" if vez in atrasadas else "coir-carimbo"
                ui.label(_r(vez.total)).classes(f"coir-numero text-3xl {classe}")
                vence = f"{vez.vencimento:%d/%m/%Y}" if vez.vencimento else ""
                juros = " (com multa e juros até hoje)" if vez.multa and vez.multa.valor else ""
                ui.label(
                    f"Competência {vez.mes:02d}/{vez.ano} · vence {vence} · "
                    f"{vez.situacao.value}{juros}"
                ).classes("text-sm")
            ui.button(
                "Ver as DARFs", icon="request_quote", on_click=lambda: ui.navigate.to("/darf")
            ).props("flat dense")
        with _cartao("Pendências").mark("cartao-pendencias"):
            anos = sorted({x.ano for x in atrasadas})
            itens = [
                (len(atrasadas), "DARF(s) em atraso"
                 + (f" ({', '.join(map(str, anos))})" if anos else ""), "/darf"),
                (len(incompletos), "arquivo(s) com nota rejeitada", "/importar"),
                (len(avisos), "aviso(s) do motor no ano", "/apuracao"),
            ]  # fmt: skip
            if not any(n for n, _, _ in itens):
                ui.label("Nenhuma pendência.").classes("coir-positivo")
            for n, texto, rota in itens:
                if n:
                    with ui.row().classes("items-center gap-2"):
                        ui.icon("error_outline", size="18px").classes("coir-carimbo")
                        ui.link(f"{n} {texto}", rota).classes("text-sm")
        with _cartao("Atalhos"):
            ui.button(
                "Importar notas", icon="upload_file", on_click=lambda: ui.navigate.to("/importar")
            ).props("outline")
            ui.button(
                "Exportar relatórios",
                icon="description",
                on_click=lambda: ui.navigate.to("/relatorios"),
            ).props("outline")
    if not meses:
        ui.label("Nada apurado neste ano: importe as notas.").classes("coir-vazio w-full")
        return
    _numeros(painel.resumo(meses, darfs), MESES[meses[-1].mes - 1].lower())
    _grafico(meses)


def _numeros(r: painel.ResumoAno, ultimo_mes: str) -> None:
    curtas = {Categoria.COMUM: "comum", Categoria.DAY_TRADE: "DT", Categoria.FII: "FII"}
    with ui.row().classes("w-full gap-8 coir-folha").mark("numeros-ano"):
        for rotulo, valor in (
            ("Resultado do ano, antes da compensação (sem o isento)",
             " · ".join(f"{curtas[c]} {_r(v)}" for c, v in r.resultado.items())),
            ("Imposto devido no ano (depois do IRRF)", _r(r.imposto_devido)),
            ("DARFs pagas", f"{r.darfs_pagas} de {r.darfs_emitidas}"),
            (f"Prejuízo a compensar em {ultimo_mes}",
             " · ".join(f"{curtas[c]} {_r(v)}" for c, v in r.prejuizo_a_compensar.items())),
        ):  # fmt: skip
            with ui.column().classes("gap-0"):
                ui.label(valor).classes("coir-numero text-lg")
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
