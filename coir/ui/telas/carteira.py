"""Tela 6 — Carteira e preço médio: a posição numa data (visão por titular, corretora ou conta)
e, ao escolher um ativo, a trilha do PM ao lado (cada passo com a regra e a nota de origem)."""

from collections.abc import Callable
from datetime import date
from decimal import Decimal

from nicegui import run, ui

from coir.domain.fiscal import Movimento, Visao
from coir.erros import CoirError
from coir.reports.modelo import Coluna, TipoColuna, Valor, formatar
from coir.servicos import carteira
from coir.servicos.estado import Estado
from coir.ui import tabela

T = TipoColuna
VISOES = {Visao.TITULAR: "Titular (fiscal)", Visao.CORRETORA: "Corretora", Visao.CONTA: "Conta"}
COLUNAS_TRILHA = (
    Coluna("Data", T.DATA),
    Coluna("C/V"),
    Coluna("Qtd.", T.QUANTIDADE),
    Coluna("Valor", T.DINHEIRO),
    Coluna("Resultado", T.DINHEIRO),
    Coluna("Qtd. após", T.QUANTIDADE),
    Coluna("PM após", T.PM),
    Coluna("Regra"),
)
ZERO = Decimal(0)


def _colunas(visao: Visao) -> tuple[Coluna, ...]:
    onde = {
        Visao.TITULAR: (),
        Visao.CORRETORA: (Coluna("Corretora"),),
        Visao.CONTA: (Coluna("Corretora"), Coluna("Conta")),
    }[visao]
    return (
        Coluna("Ativo"),
        Coluna("Tipo"),
        *onde,
        Coluna("Quantidade", T.QUANTIDADE),
        Coluna("PM", T.PM),
        Coluna("Custo", T.DINHEIRO),
    )


def _valores(linhas: list[carteira.LinhaPosicao], visao: Visao) -> list[tuple[Valor, ...]]:
    def onde(p: carteira.LinhaPosicao) -> tuple[Valor, ...]:
        if visao is Visao.TITULAR:
            return ()
        return (p.corretora,) if visao is Visao.CORRETORA else (p.corretora, p.numero_conta)

    return [(p.ticker, p.tipo, *onde(p), p.quantidade, p.pm, p.custo) for p in linhas]


def _total(linhas: list[carteira.LinhaPosicao], visao: Visao) -> tuple[Valor, ...]:
    compradas = sum((p.custo for p in linhas if p.custo > 0), ZERO)
    vazias = {Visao.TITULAR: 0, Visao.CORRETORA: 1, Visao.CONTA: 2}[visao]
    return (f"{len(linhas)} posições", "", *([""] * vazias), None, None, compradas)


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Carteira e preço médio").classes("coir-titulo-pagina")
    ui.label(
        "A posição numa data e a trilha do preço médio de cada ativo. O PM fiscal é o da visão "
        "por titular: um custo médio por ativo, somando todas as corretoras e contas (IN RFB "
        "1.585/2015 art. 58). As visões por corretora e conta servem para conferir a custódia."
    ).classes("coir-subtitulo")
    escolha: dict[str, object] = {"data": estado.fim_do_ano(), "visao": Visao.TITULAR}

    with ui.row().classes("w-full items-center gap-3"):
        inicial = escolha["data"]
        assert isinstance(inicial, date)
        campo_data = (
            ui.input("Posição em", value=inicial.isoformat())
            .props("outlined dense type=date")
            .classes("w-[170px]")
            .mark("data-carteira")
        )
        ui.select(
            {v: r for v, r in VISOES.items()},
            value=Visao.TITULAR,
            label="Visão",
            on_change=lambda e: mudar("visao", Visao(e.value)),
        ).props("outlined dense options-dense").classes("w-[190px]").mark("visao")
        resumo = ui.label("").classes("coir-subtitulo text-sm")
    campo_data.on_value_change(lambda e: mudar("data", e.value))

    with ui.row().classes("w-full items-start gap-4"):
        lista = ui.column().classes("min-w-[440px] flex-[0.8] gap-2")
        lado = ui.column().classes("min-w-[520px] flex-[1.2] gap-2")
    linhas: list[carteira.LinhaPosicao] = []

    async def carregar() -> None:
        ate, visao = escolha["data"], escolha["visao"]
        assert isinstance(ate, date) and isinstance(visao, Visao)
        lista.clear()
        with lista:
            ui.spinner(size="lg")
        try:
            resultado = await run.io_bound(estado.carteira, ate, visao)
        except CoirError as erro:
            lista.clear()
            with lista:
                ui.label(f"Não foi possível calcular a carteira: {erro}").classes("coir-negativo")
            return
        if resultado is None:
            return
        posicoes, contas = resultado
        linhas[:] = carteira.posicoes(posicoes, contas, estado.titular)
        lista.clear()
        with lista:
            if not linhas:
                ui.label("Nenhuma posição aberta nesta data.").classes("coir-vazio w-full")
            else:
                tabela.grade(
                    _colunas(visao),
                    _valores(linhas, visao),
                    chaves=list(range(len(linhas))),
                    total=_total(linhas, visao),
                    altura="h-[620px]",
                    ao_clicar=lambda i: abrir(int(str(i))),
                ).mark("grade-carteira")
        resumo.set_text(f"{len(linhas)} posições em {ate:%d/%m/%Y}")
        lado.clear()
        with lado:
            ui.label("Escolha um ativo para ver a trilha do preço médio.").classes(
                "coir-vazio w-full"
            )

    async def abrir(indice: int) -> None:
        p = linhas[indice]
        ate, visao = escolha["data"], escolha["visao"]
        assert isinstance(ate, date) and isinstance(visao, Visao)
        movs = await run.io_bound(estado.movimentos, ate, visao)
        mapa = await run.io_bound(estado.notas_das_operacoes)
        if movs is None or mapa is None:
            return
        passos = carteira.trilha(movs, estado.titular, p, ate)
        lado.clear()
        with lado:
            _trilha(p, passos, mapa)

    def mudar(chave: str, valor: object) -> None:
        if chave == "data":
            try:
                valor = date.fromisoformat(str(valor))
            except ValueError:
                return
        escolha[chave] = valor
        ui.timer(0.01, carregar, once=True)

    ui.timer(0.01, carregar, once=True)


def _trilha(p: carteira.LinhaPosicao, passos: list[Movimento], mapa: dict[int, int]) -> None:
    onde = " · ".join(x for x in (p.corretora, p.numero_conta and f"conta {p.numero_conta}") if x)
    with ui.row().classes("items-baseline gap-3"):
        ui.label(f"Trilha do PM: {p.ticker}").classes("coir-titulo text-lg")
        if onde:
            ui.label(onde).classes("coir-subtitulo text-sm")
    ui.label(
        f"Posição: {formatar(p.quantidade, T.QUANTIDADE)} · PM {formatar(p.pm, T.PM)} · "
        f"custo {formatar(p.custo, T.DINHEIRO)}"
    ).classes("coir-numero")
    valores: list[tuple[Valor, ...]] = [
        (
            m.data,
            m.cv or "",
            m.quantidade,
            m.valor,
            m.resultado or None,
            m.quantidade_apos,
            m.pm_apos,
            m.regra,
        )
        for m in passos
    ]
    notas = [carteira.nota_de(m.origem, mapa) for m in passos]
    tabela.grade(
        COLUNAS_TRILHA,
        valores,
        chaves=notas,
        altura="h-[560px]",
        pelo_conteudo=True,
        ao_clicar=lambda n: n is not None and ui.navigate.to(f"/notas/{n}"),
    ).mark("trilha")
    ui.label("Clique num passo para abrir a nota de origem (eventos não têm nota).").classes(
        "coir-subtitulo text-xs"
    )
