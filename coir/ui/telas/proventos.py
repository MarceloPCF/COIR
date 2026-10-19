"""Tela 10 — Proventos do ano: cartões pelas linhas do IRPF, gráfico mensal e a tabela dos
lançamentos (da movimentação da B3), com os avisos (Lei 15.270, BDR, a classificar)."""

from collections.abc import Callable
from decimal import Decimal

from nicegui import run, ui

from coir.domain.proventos import Provento, TipoProvento
from coir.erros import CoirError
from coir.reports.fiscais import MESES, NOME_PROVENTO
from coir.reports.modelo import Coluna, TipoColuna, Valor, formatar
from coir.servicos.estado import Estado
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

    ui.timer(0.01, carregar, once=True)


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
