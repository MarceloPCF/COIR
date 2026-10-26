"""Tela 9 — Futuros: as posições abertas (ajustes acumulados ainda não realizados), o resultado do
ano por mês e contrato e, ao escolher um contrato, os ajustes dia a dia com as notas."""

from collections.abc import Callable
from decimal import Decimal

from nicegui import run, ui

from coir.domain.fiscal import Grupo
from coir.erros import CoirError
from coir.reports.fiscais import MESES
from coir.reports.modelo import Coluna, TipoColuna, Valor
from coir.servicos.derivativos import Futuros, resultado_mensal
from coir.servicos.estado import Estado
from coir.ui import tabela

T = TipoColuna
ZERO = Decimal(0)
GRUPOS = {Grupo.INDICE: "Índice", Grupo.DOLAR: "Dólar", Grupo.JUROS: "Juros",
          Grupo.COMMODITY: "Commodity"}  # fmt: skip
COLUNAS_ABERTAS = (
    Coluna("Contrato"),
    Coluna("Grupo"),
    Coluna("Contratos", T.QUANTIDADE),
    Coluna("Ajustes acumulados", T.DINHEIRO),
    Coluna("Vencimento", T.DATA),
)
COLUNAS_MES = (
    Coluna("Mês"),
    Coluna("Contrato"),
    Coluna("Day trade", T.DINHEIRO),
    Coluna("Comum", T.DINHEIRO),
    Coluna("Custos", T.DINHEIRO),
)
COLUNAS_AJUSTES = (
    Coluna("Data", T.DATA),
    Coluna("Nota"),
    Coluna("Tipo"),
    Coluna("C/V"),
    Coluna("Contratos", T.QUANTIDADE),
    Coluna("Ajuste", T.DINHEIRO),
)


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Futuros").classes("coir-titulo-pagina")
    ui.label(
        "O resultado de cada contrato sai no encerramento ou no vencimento: a soma dos ajustes "
        "diários desde a abertura, menos os custos (Lei 11.051/2004 art. 32). Enquanto a posição "
        "está aberta, os ajustes ficam acumulados. Day trade tem resultado no dia."
    ).classes("coir-subtitulo")
    corpo = ui.column().classes("w-full gap-3")

    async def carregar() -> None:
        with corpo:
            espera = ui.spinner(size="lg")
        try:
            dados = await run.io_bound(estado.futuros)
        except CoirError as erro:
            corpo.clear()
            with corpo:
                ui.label(f"Não foi possível calcular: {erro}").classes("coir-negativo")
            return
        espera.delete()
        with corpo:
            if dados is None or not (dados.abertas or dados.realizados):
                ui.label("Nenhuma operação com futuros até esta data.").classes("coir-vazio w-full")
                return
            _conteudo(dados, estado.ano)

    ui.timer(0.01, carregar, once=True)


def _conteudo(dados: Futuros, ano: int | None) -> None:
    def abrir(contrato: object) -> None:
        linhas = [a for a in dados.ajustes if a.mercadoria == contrato]
        ajustes.clear()
        with ajustes:
            ui.label(f"Ajustes de {contrato}").classes("coir-titulo text-lg")
            valores: list[tuple[Valor, ...]] = [
                (a.data, a.numero, a.tipo, a.cv, a.quantidade or None, a.valor) for a in linhas
            ]
            tabela.grade(
                COLUNAS_AJUSTES,
                valores,
                chaves=[a.nota for a in linhas],
                total=(
                    "",
                    "",
                    f"{len(linhas)} linhas",
                    "",
                    None,
                    sum((a.valor for a in linhas), ZERO),
                ),
                altura="h-[420px]",
                ao_clicar=lambda n: ui.navigate.to(f"/notas/{n}"),
            ).mark("grade-ajustes")
            ui.label("Clique numa linha para abrir a nota.").classes("coir-subtitulo text-xs")

    ui.label("Posições abertas").classes("coir-titulo text-base")
    if dados.abertas:
        tabela.grade(
            COLUNAS_ABERTAS,
            [
                (
                    p.mercadoria,
                    GRUPOS.get(p.grupo, p.grupo.value),
                    p.quantidade,
                    p.acumulado,
                    p.vencimento,
                )
                for p in dados.abertas
            ],
            chaves=[p.mercadoria for p in dados.abertas],
            ao_clicar=abrir,
        ).mark("grade-futuros-abertos")
    else:
        ui.label("Nenhum contrato em aberto.").classes("coir-vazio w-full")
    linhas = resultado_mensal(dados.realizados, ano) if ano else []
    ui.label(f"Resultado de {ano}").classes("coir-titulo text-base")
    if linhas:
        tabela.grade(
            COLUNAS_MES,
            [
                (MESES[m - 1], c, dt or None, comum or None, custos)
                for m, c, dt, comum, custos in linhas
            ],
            chaves=[c for _, c, *_ in linhas],
            total=(
                "Total",
                f"{len({c for _, c, *_ in linhas})} contratos",
                sum((x[2] for x in linhas), ZERO),
                sum((x[3] for x in linhas), ZERO),
                sum((x[4] for x in linhas), ZERO),
            ),
            altura="h-[420px]",
            ao_clicar=abrir,
        ).mark("grade-futuros-mes")
    else:
        ui.label("Nenhum resultado realizado neste ano.").classes("coir-vazio w-full")
    for aviso in dados.avisos:
        ui.label(aviso).classes("coir-carimbo text-sm")
    ajustes = ui.column().classes("w-full gap-2")  # abrir() a preenche
    with ajustes:
        ui.label("Escolha um contrato para ver os ajustes dia a dia.").classes("coir-vazio w-full")
