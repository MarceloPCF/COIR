"""Tela 5 — Operações: o livro de operações do ano numa grade, com filtros e totais."""

from collections.abc import Callable
from decimal import Decimal

from nicegui import ui
from nicegui.events import ValueChangeEventArguments

from coir.reports.fiscais import MESES
from coir.reports.modelo import Coluna, TipoColuna, Valor, formatar
from coir.servicos.estado import Estado
from coir.servicos.operacoes import LinhaOperacao
from coir.ui import tabela

T = TipoColuna
COLUNAS = (
    Coluna("Data", T.DATA),
    Coluna("Corretora"),
    Coluna("Conta"),
    Coluna("Nota"),
    Coluna("Ativo"),
    Coluna("Mercado"),
    Coluna("C/V"),
    Coluna("Tipo"),
    Coluna("Quantidade", T.QUANTIDADE),
    Coluna("Preço", T.PRECO),
    Coluna("Total", T.DINHEIRO),
    Coluna("Custos", T.DINHEIRO),
)
ZERO = Decimal(0)


def _valores(linhas: list[LinhaOperacao]) -> list[tuple[Valor, ...]]:
    return [
        (
            o.data,
            o.corretora,
            o.conta,
            o.nota,
            o.ativo,
            o.mercado,
            o.cv,
            "Day trade" if o.day_trade else "Comum",
            o.quantidade,
            o.preco,
            o.total,
            o.custos,
        )
        for o in linhas
    ]


def _total(linhas: list[LinhaOperacao]) -> tuple[Valor, ...]:
    compras = sum((o.total for o in linhas if o.cv == "C"), ZERO)
    vendas = sum((o.total for o in linhas if o.cv == "V"), ZERO)
    custos = sum((o.custos or ZERO for o in linhas), ZERO)
    return (
        f"{len(linhas)} operações",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        None,
        None,
        f"C {formatar(compras, T.DINHEIRO)} · V {formatar(vendas, T.DINHEIRO)}",
        custos,
    )


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Operações").classes("coir-titulo-pagina")
    ui.label(
        f"O livro de operações de {estado.ano}, como lido das notas. Custos: os da nota "
        "rateados pelo valor de cada operação (no BM&F, entram no resultado do dia). Clique "
        'numa operação para abrir a nota. Para o histórico de um ativo, ligue "Todos os anos" '
        "e digite o código.",
    ).classes("coir-subtitulo")
    todas = estado.livro()
    if not todas and not estado.livro(todos_os_anos=True):
        ui.label("Nenhuma operação importada.").classes("coir-vazio w-full")
        return
    filtro: dict[str, object] = {"ativo": "", "corretora": "", "mercado": "", "mes": 0, "dt": False}

    def ano_todo(todos: bool) -> None:
        todas[:] = estado.livro(todos_os_anos=todos)
        mudar("mes", 0 if todos else filtro["mes"])

    def vistas() -> list[LinhaOperacao]:
        ativo = str(filtro["ativo"]).strip().upper()
        return [
            o
            for o in todas
            if ativo in o.ativo
            and (not filtro["corretora"] or o.corretora == filtro["corretora"])
            and (not filtro["mercado"] or o.mercado == filtro["mercado"])
            and (not filtro["mes"] or o.data.month == filtro["mes"])
            and (not filtro["dt"] or o.day_trade)
        ]

    def mudar(chave: str, valor: object) -> None:
        filtro[chave] = valor
        linhas = vistas()
        tabela.atualizar(
            grade,
            COLUNAS,
            _valores(linhas),
            chaves=[o.id_nota for o in linhas],
            total=_total(linhas),
        )

    def ao_escolher(chave: str) -> Callable[[ValueChangeEventArguments], None]:
        return lambda e: mudar(chave, e.value)

    with ui.row().classes("w-full items-center gap-3"):
        ui.input("Ativo", on_change=lambda e: mudar("ativo", e.value or "")).props(
            "outlined dense clearable"
        ).classes("w-[150px]").mark("filtro-ativo")
        for chave, rotulo, opcoes, largura in (
            ("corretora", "Corretora", sorted({o.corretora for o in todas}), "w-[150px]"),
            ("mercado", "Mercado", sorted({o.mercado for o in todas}), "w-[170px]"),
        ):
            ui.select(
                {"": "Todos", **{x: x for x in opcoes}},
                value="",
                label=rotulo,
                on_change=ao_escolher(chave),
            ).props("outlined dense options-dense").classes(largura)
        ui.select(
            {0: "Todos os meses", **{i + 1: m for i, m in enumerate(MESES)}},
            value=0,
            label="Mês",
            on_change=lambda e: mudar("mes", e.value),
        ).props("outlined dense options-dense").classes("w-[170px]")
        ui.checkbox("Só day trade", on_change=lambda e: mudar("dt", e.value)).mark("filtro-dt")
        ui.switch("Todos os anos", on_change=lambda e: ano_todo(bool(e.value))).mark(
            "filtro-todos-os-anos")  # fmt: skip
    grade = tabela.grade(
        COLUNAS,
        _valores(todas),
        chaves=[o.id_nota for o in todas],
        total=_total(todas),
        altura="h-[680px]",
        ao_clicar=lambda nota: ui.navigate.to(f"/notas/{nota}"),
    )
    grade.mark("grade-operacoes")
