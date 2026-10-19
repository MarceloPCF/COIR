"""Tela 16 — Cadastros, em abas: titulares (adicionar e renomear), contas e corretoras, ativos e
tickers, e os parâmetros fiscais com vigência e fonte (consulta: vêm da importação e de data/)."""

from collections.abc import Callable
from datetime import date

from nicegui import run, ui

from coir.domain.cpf import cpf_valido
from coir.erros import CoirError
from coir.reports.modelo import Coluna, TipoColuna, Valor
from coir.servicos import cadastros
from coir.servicos.estado import Estado
from coir.ui import tabela

T = TipoColuna
COLUNAS_CONTAS = (
    Coluna("Corretora"),
    Coluna("Conta"),
    Coluna("Titular"),
    Coluna("Notas", T.INTEIRO),
    Coluna("Primeira nota", T.DATA),
    Coluna("Última nota", T.DATA),
)
COLUNAS_TICKERS = (
    Coluna("Ticker"),
    Coluna("Ativo", T.INTEIRO),
    Coluna("Desde", T.DATA),
    Coluna("Até", T.DATA),
    Coluna("ISIN"),
)
COLUNAS_PARAMETROS = (
    Coluna("Parâmetro"),
    Coluna("Valor"),
    Coluna("Início", T.DATA),
    Coluna("Fim", T.DATA),
    Coluna("Fonte"),
)


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Cadastros").classes("coir-titulo-pagina")
    ui.label(
        "Titulares, contas, ativos e parâmetros fiscais. Contas, corretoras e ativos nascem da "
        "importação das notas; tickers e parâmetros vêm das tabelas de referência (data/)."
    ).classes("coir-subtitulo")
    with ui.tabs().props("no-caps align=left").classes("w-full") as abas:
        t_titulares = ui.tab("Titulares").mark("aba-titulares")
        t_contas = ui.tab("Contas e corretoras").mark("aba-contas")
        t_ativos = ui.tab("Ativos e tickers").mark("aba-ativos")
        t_parametros = ui.tab("Parâmetros fiscais").mark("aba-parametros")
    with ui.tab_panels(abas, value=t_titulares).classes("w-full bg-transparent"):
        with ui.tab_panel(t_titulares).classes("p-0 pt-3"):
            _titulares(estado, ao_mudar)
        with ui.tab_panel(t_contas).classes("p-0 pt-3"):
            contas = cadastros.contas(estado.banco) if estado.existe else []
            valores: list[tuple[Valor, ...]] = [
                (c.corretora, c.numero, c.titular, c.notas, c.primeira, c.ultima) for c in contas
            ]
            tabela.grade(COLUNAS_CONTAS, valores, altura="h-[560px]").mark("grade-contas")
            ui.label(
                "A corretora vem do código impresso na nota; Clear usa o perfil da XP e Necton "
                "o da BTG."
            ).classes("coir-subtitulo text-xs")
        with ui.tab_panel(t_ativos).classes("p-0 pt-3"):
            tickers = cadastros.tickers(estado.banco) if estado.existe else []
            valores = [(x.ticker, x.ativo, x.inicio, x.fim, x.isin) for x in tickers]
            busca = ui.input("Buscar ticker").props("outlined dense clearable").classes("w-[220px]")
            grade = tabela.grade(COLUNAS_TICKERS, valores, altura="h-[560px]")
            grade.mark("grade-tickers")

            def filtrar(texto: str) -> None:
                vistos = [v for v in valores if texto.strip().upper() in str(v[0])]
                tabela.atualizar(grade, COLUNAS_TICKERS, vistos)

            busca.on_value_change(lambda e: filtrar(e.value or ""))
            ui.label(
                "Trocas de ticker (data/trocas_ticker.csv) mantêm o mesmo ativo e o mesmo PM."
            ).classes("coir-subtitulo text-xs")
        with ui.tab_panel(t_parametros).classes("p-0 pt-3"):
            parametros = sorted(
                estado.fiscais.parametros.parametros, key=lambda p: (p.nome, p.inicio or date.min)
            )
            valores = [(p.nome, p.valor, p.inicio, p.fim, p.fonte) for p in parametros]
            tabela.grade(COLUNAS_PARAMETROS, valores, altura="h-[560px]", pelo_conteudo=True).mark(
                "grade-parametros"
            )
            ui.label(
                "Nenhum parâmetro fiscal fica no código: alíquotas, limites, códigos e "
                "prazos vêm daqui, com vigência (data/parametros_fiscais.csv)."
            ).classes("coir-subtitulo text-xs")


def _titulares(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    @ui.refreshable
    def lista() -> None:
        titulares = cadastros.titulares(estado.banco) if estado.existe else []
        if not titulares:
            ui.label("Nenhum titular cadastrado.").classes("coir-vazio w-full")
        for t in titulares:
            with ui.row().classes("coir-folha w-full items-start gap-4"):
                with ui.column().classes("gap-0 grow"):
                    ui.label(t.nome).classes("coir-titulo text-base").mark(f"titular-{t.id}")
                    ui.label(f"CPF {t.cpf} · id {t.id}").classes("coir-subtitulo text-sm")
                    contas = ", ".join(t.contas) or "sem contas (nenhuma nota importada)"
                    ui.label(f"Contas: {contas}").classes("text-sm")
                ui.button("Renomear", icon="edit", on_click=_renomear(estado, t, recarregar)).props(
                    "flat dense"
                ).mark(f"renomear-{t.id}")

    def recarregar() -> None:
        estado.invalidar()
        lista.refresh()
        ao_mudar()

    lista()
    with (
        ui.expansion("Novo titular", icon="person_add")
        .classes("coir-folha w-full")
        .mark("novo-titular")
    ):
        nome = ui.input("Nome").props("outlined dense").classes("w-[360px]").mark("novo-nome")
        cpf = (
            ui.input("CPF", validation={"CPF inválido": lambda v: cpf_valido(v or "")})
            .props("outlined dense mask='###.###.###-##' lazy-rules")
            .classes("w-[220px]")
            .mark("novo-cpf")
        )

        async def adicionar() -> None:
            try:
                await run.io_bound(
                    cadastros.adicionar_titular, estado.banco, cpf.value or "", nome.value or ""
                )
            except CoirError as erro:
                ui.notify(str(erro), type="negative")
                return
            nome.value = cpf.value = ""
            ui.notify("Titular cadastrado.", type="positive")
            recarregar()

        ui.button("Cadastrar", icon="save", on_click=adicionar).mark("cadastrar-titular")
        ui.label(
            "Notas de um CPF ainda não cadastrado criam o titular sozinhas, com o nome "
            "impresso; renomeie aqui se quiser."
        ).classes("coir-subtitulo text-xs")


def _renomear(
    estado: Estado, t: cadastros.TitularCadastro, recarregar: Callable[[], None]
) -> Callable[[], None]:
    def abrir() -> None:
        with ui.dialog() as janela, ui.card().classes("min-w-[420px]"):
            ui.label(f"Renomear {t.nome}").classes("coir-titulo text-base")
            campo = (
                ui.input("Nome", value=t.nome)
                .props("outlined dense autofocus")
                .classes("w-full")
                .mark("nome-novo")
            )

            def salvar() -> None:
                try:
                    cadastros.renomear_titular(estado.banco, t.id, campo.value or "")
                except CoirError as erro:
                    ui.notify(str(erro), type="negative")
                    return
                janela.close()
                recarregar()

            with ui.row().classes("w-full justify-end"):
                ui.button("Cancelar", on_click=janela.close).props("flat")
                ui.button("Salvar", on_click=salvar).mark("salvar-nome")
        janela.open()

    return abrir
