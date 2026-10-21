"""Tela 10 — Proventos do ano: cartões pelas linhas do IRPF, gráfico mensal e a tabela dos
lançamentos (da movimentação da B3), com os avisos (Lei 15.270, BDR, a classificar)."""

from collections.abc import Awaitable, Callable
from datetime import datetime
from decimal import Decimal

from nicegui import run, ui

from coir.domain.proventos import Provento, TipoProvento
from coir.erros import CoirError
from coir.reports.fiscais import MESES, NOME_PROVENTO
from coir.reports.modelo import Coluna, TipoColuna, Valor, formatar
from coir.servicos.estado import Estado
from coir.servicos.proventos import informar_credito
from coir.ui import tabela, tema

T = TipoColuna
ZERO = Decimal(0)
COLUNAS = (
    Coluna("Pagamento", T.DATA),
    Coluna("Ativo"),
    Coluna("Tipo"),
    Coluna("Instituição"),
    Coluna("Quantidade", T.QUANTIDADE),
    Coluna("Líquido", T.DINHEIRO),
    Coluna("Bruto", T.DINHEIRO),
    Coluna("IRRF", T.DINHEIRO),
    Coluna("Observação"),
)
CARTOES = (
    (TipoProvento.DIVIDENDO, "Dividendos · isentos (09)"),
    (TipoProvento.JCP, "JCP · exclusiva (10)"),
    (TipoProvento.RENDIMENTO_FII, "Rendimentos de FII · isentos"),
    (TipoProvento.REEMBOLSO_BTC, "Reembolsos de BTC"),
    (TipoProvento.ATUALIZACAO, "Atualização (Selic) · exclusiva"),
)
SERIES = (
    (TipoProvento.DIVIDENDO, tema.CLARO.caneta),
    (TipoProvento.JCP, tema.CLARO.carimbo),
    (TipoProvento.RENDIMENTO_FII, tema.CLARO.positivo),
    (TipoProvento.REEMBOLSO_BTC, tema.CLARO.tinta_suave),
)


def _r(v: Decimal) -> str:
    return formatar(v, T.DINHEIRO)


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Proventos").classes("coir-titulo-pagina")
    ui.label(
        f"Os proventos de {estado.ano} pela movimentação da B3 (Importar > Planilhas da B3), "
        "separados como no IRPF. O JCP creditado num ano e pago no seguinte entra no ano do "
        "crédito. Bruto e IRRF do JCP são estimados: vale o informe de rendimentos da empresa."
    ).classes("coir-subtitulo")
    corpo = ui.column().classes("w-full gap-4")

    async def carregar() -> None:
        with corpo:
            espera = ui.spinner(size="lg")
        try:
            resultado = await run.io_bound(estado.proventos)
        except CoirError as erro:
            corpo.clear()
            with corpo:
                ui.label(f"Não foi possível calcular: {erro}").classes("coir-negativo")
            return
        espera.delete()
        if resultado is None:
            return
        todos, avisos = resultado
        do_ano = [p for p in todos if p.ano_declaracao == estado.ano]
        with corpo:
            if not do_ano:
                ui.label("Nenhum provento neste ano. Importe a movimentação da B3 em Importar.",
                         ).classes("coir-vazio w-full")  # fmt: skip
            else:
                _conteudo(do_ano)
            for aviso in (a for a in avisos if str(estado.ano) in a[:10]):
                ui.label(aviso).classes("coir-carimbo text-sm")
            jcp = [p for p in todos if p.tipo is TipoProvento.JCP and estado.ano in (
                p.ano_declaracao, p.data.year)]  # fmt: skip
            if jcp:
                _creditos(estado, jcp, recarregar)

    async def recarregar() -> None:
        corpo.clear()
        ao_mudar()
        await carregar()

    ui.timer(0.01, carregar, once=True)


def _creditos(estado: Estado, jcp: list[Provento],
              recarregar: Callable[[], Awaitable[None]]) -> None:  # fmt: skip
    """Data do crédito de cada JCP: a da B3 (a mais tarde entre a aprovação e a data com) ou a
    informada pelo usuário, que decide o ano da declaração (IN SRF 41/1998 art. 3º)."""

    async def salvar(p: Provento, campo: ui.input, apagar: bool) -> None:
        try:
            texto = (campo.value or "").strip()
            dia = None if apagar else datetime.strptime(texto, "%d/%m/%Y").date()
        except ValueError:
            ui.notify("Data inválida: use dd/mm/aaaa.", type="warning")
            return
        try:
            await run.io_bound(informar_credito, estado.banco, p.chave, dia, estado.agora())
        except CoirError as erro:
            ui.notify(f"Não gravado: {erro}", type="negative")
            return
        ui.notify("Data do crédito gravada." if dia else "Data informada apagada.",
                  type="positive")  # fmt: skip
        await recarregar()

    def clique(p: Provento, campo: ui.input, apagar: bool = False
               ) -> Callable[[], Awaitable[None]]:  # fmt: skip
        return lambda: salvar(p, campo, apagar)

    async def propor() -> None:
        try:
            resultado = await run.io_bound(estado.propor_creditos_dirpf)
        except CoirError as erro:
            ui.notify(f"Não foi possível ler as declarações: {erro}", type="negative")
            return
        if resultado is None:
            return
        propostas, avisos = resultado
        for aviso in avisos:
            ui.notify(aviso, type="warning", multi_line=True)
        if not propostas:
            ui.notify("Nenhum JCP a propor pelas declarações em _dirpf (99.07).", type="info")
            return
        with ui.dialog() as janela, ui.card().classes("min-w-[560px]").mark("propostas-credito"):
            ui.label("Créditos propostos pela DIRPF entregue").classes("coir-titulo text-base")
            ui.label(
                "JCP a receber em 31/12 (Bens e Direitos 99.07) cuja soma fecha com os JCP pagos "
                "no ano seguinte pela mesma empresa: creditados no ano anterior. A data fica "
                "31/12 do ano do crédito (a declaração só diz o ano)."
            ).classes("coir-subtitulo text-xs")
            caixas = [
                (ui.checkbox(f"{x.pagamento:%d/%m/%Y} {x.codigo} {_r(x.liquido)} -> creditado em "
                             f"{x.credito.year}", value=True), x)
                for x in propostas
            ]  # fmt: skip

            async def confirmar() -> None:
                escolhidas = [x for caixa, x in caixas if caixa.value]
                try:
                    await run.io_bound(estado.confirmar_creditos, escolhidas)
                except CoirError as erro:
                    ui.notify(f"Não gravado: {erro}", type="negative")
                    return
                janela.close()
                ui.notify(f"{len(escolhidas)} data(s) de crédito gravada(s).", type="positive")
                await recarregar()

            with ui.row().classes("justify-end w-full"):
                ui.button("Cancelar", on_click=janela.close).props("flat")
                ui.button("Confirmar", on_click=confirmar).mark("confirmar-propostas")
        janela.open()

    with ui.column().classes("coir-folha w-full gap-2").mark("creditos-jcp"):
        with ui.row().classes("w-full items-center"):
            ui.label("Data do crédito do JCP").classes("coir-titulo text-base")
            ui.space()
            ui.button("Propor pela DIRPF", icon="auto_fix_high", on_click=propor).props(
                "flat dense").mark("propor-dirpf")  # fmt: skip
        ui.label(
            "O JCP é declarado no ano do crédito. A data vem do pacote público da B3 (aprovação) "
            "quando ele cobre o pagamento; sem ela, o COIR só percebe o crédito no ano anterior "
            "pela mudança de alíquota. Corrija pelo informe de rendimentos da empresa, ou use a "
            "DIRPF entregue (99.07) em _dirpf na pasta de dados."
        ).classes("coir-subtitulo text-xs")
        for p in sorted(jcp, key=lambda p: (p.data, p.codigo)):
            with ui.row().classes("items-center gap-3 w-full"):
                ui.label(f"{p.data:%d/%m/%Y} {p.codigo}").classes("coir-numero w-[180px]")
                ui.label(_r(p.liquido)).classes("coir-numero w-[110px] text-right")
                origem = p.credito_origem or "sem data"
                ui.label(f"declarado em {p.ano_declaracao} · {origem}").classes(
                    "coir-subtitulo text-xs w-[200px]")  # fmt: skip
                campo = ui.input("Creditado em", placeholder="dd/mm/aaaa",
                                 value=f"{p.credito:%d/%m/%Y}" if p.credito else "",
                                 ).props("outlined dense").classes("w-[150px]").mark(
                    f"credito-{p.chave[:12]}")  # fmt: skip
                ui.button("Salvar", on_click=clique(p, campo)).props("flat dense").mark(
                    f"salvar-credito-{p.chave[:12]}")  # fmt: skip
                if origem == "informada":
                    ui.button("Limpar", on_click=clique(p, campo, apagar=True)).props(
                        "flat dense")  # fmt: skip


def _conteudo(do_ano: list[Provento]) -> None:
    with ui.row().classes("w-full items-stretch gap-3").mark("cartoes-proventos"):
        for tipo, rotulo in CARTOES:
            lista = [p for p in do_ano if p.tipo is tipo]
            if not lista and tipo not in (TipoProvento.DIVIDENDO, TipoProvento.JCP):
                continue
            with ui.column().classes("coir-folha gap-0 min-w-[190px] flex-1"):
                ui.label(rotulo).classes("coir-campo-rotulo")
                ui.label(_r(sum((p.liquido for p in lista), ZERO))).classes("coir-numero text-2xl")
                if tipo in (TipoProvento.JCP, TipoProvento.ATUALIZACAO) and lista:
                    irrf = sum((p.irrf for p in lista), ZERO)
                    ui.label(f"IRRF {_r(irrf)} (estimado)").classes("coir-subtitulo text-xs")
        a_receber = [p for p in do_ano if p.a_receber_em_31_12]
        if a_receber:
            with ui.column().classes("coir-folha gap-0 min-w-[190px] flex-1"):
                ui.label("JCP a receber em 31/12 (99.07)").classes("coir-campo-rotulo")
                ui.label(_r(sum((p.liquido for p in a_receber), ZERO))).classes(
                    "coir-numero text-2xl")  # fmt: skip
    meses = sorted({(p.data.year, p.data.month) for p in do_ano})
    tinta = "#7A8792"
    ui.echart({
        "animation": False,
        "grid": {"left": 70, "right": 20, "top": 40, "bottom": 30},
        "legend": {"top": 0, "textStyle": {"color": tinta}},
        "tooltip": {"trigger": "axis", ":valueFormatter":
                    "v => v.toLocaleString('pt-BR', {minimumFractionDigits: 2})"},
        "xAxis": {"type": "category", "axisLabel": {"color": tinta},
                  "data": [f"{MESES[m - 1][:3]}/{a % 100:02d}" for a, m in meses]},
        "yAxis": {"type": "value", "axisLabel": {"color": tinta}},
        "series": [
            {"name": NOME_PROVENTO[tipo], "type": "bar", "stack": "total",
             "itemStyle": {"color": cor},
             "data": [float(sum((p.liquido for p in do_ano if p.tipo is tipo
                                 and (p.data.year, p.data.month) == am), ZERO))  # só desenho
                      for am in meses]}
            for tipo, cor in SERIES
        ],
    }).classes("w-full h-[300px] coir-folha").mark("grafico-proventos")  # fmt: skip
    valores: list[tuple[Valor, ...]] = [
        (p.data, p.codigo, NOME_PROVENTO[p.tipo], p.instituicao, p.quantidade, p.liquido, p.bruto,
         p.irrf, p.observacao)
        for p in sorted(do_ano, key=lambda p: (p.data, p.codigo))
    ]  # fmt: skip
    tabela.grade(COLUNAS, valores, altura="h-[480px]", pelo_conteudo=True).mark("grade-proventos")
