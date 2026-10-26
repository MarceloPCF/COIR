"""Tela 16 — Cadastros, em abas: titulares (adicionar e renomear), contas e corretoras, ativos e
tickers, e os parâmetros fiscais com vigência e fonte (consulta: vêm da importação e de data/)."""

from collections.abc import Callable
from datetime import date
from decimal import Decimal

from nicegui import run, ui

from coir.domain.cpf import cpf_valido
from coir.domain.fiscal import SaldoInicial, TipoEvento
from coir.erros import CoirError
from coir.reports.modelo import Coluna, TipoColuna, Valor, formatar
from coir.servicos import cadastros, eventos, saldos_iniciais
from coir.servicos.estado import Estado
from coir.ui import tabela
from coir.ui.entrada import valor_br

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


def _saldos(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    """Para quem começa sem as notas antigas: o prejuízo a compensar em 31/12 (aqui) e a posição
    de cada ativo em 31/12 (eventos "Custódia inicial", lançados aqui em grupo)."""
    ui.label(
        "Se você não tem todas as notas desde o começo, informe a situação em 31/12 de um ano: o "
        "prejuízo a compensar (ficha de renda variável de dezembro da última declaração: "
        "operações comuns, day trade e FII, cada um separado) e, abaixo, a custódia de cada ativo "
        "(quantidade e custo do Bens e Direitos). Importe as notas a partir de janeiro "
        "seguinte."
    ).classes("coir-subtitulo text-sm")
    if not estado.existe or estado.titular is None:
        ui.label("Escolha ou cadastre um titular primeiro.").classes("coir-vazio w-full")
        return
    titular = estado.titular
    atual = saldos_iniciais.ler(estado.banco, titular)
    hoje = estado.agora().date()
    with ui.column().classes("coir-folha gap-2 w-full max-w-[640px]").mark("form-saldos"):
        ano = ui.select(list(range(hoje.year - 1, 2000, -1)), label="Situação em 31/12 de",
                        value=atual.ano if atual else hoje.year - 1)  # fmt: skip
        ano.props("outlined dense").classes("w-[200px]").mark("saldo-ano")
        campos = {}
        for chave, rotulo in (("comum", "Prejuízo a compensar: operações comuns"),
                              ("day_trade", "Prejuízo a compensar: day trade"),
                              ("fii", "Prejuízo a compensar: FII e Fiagro")):  # fmt: skip
            valor = getattr(atual, chave) if atual else None
            campos[chave] = ui.input(rotulo, value=_moeda(valor) if valor else "").props(
                "outlined dense input-class=text-right").classes("w-[360px]").mark(
                f"saldo-{chave}")  # fmt: skip

        async def salvar(apagar: bool = False) -> None:
            saldo = None
            if not apagar:
                valores: dict[str, Decimal] = {}
                for chave, campo in campos.items():
                    valor = valor_br(campo.value or "0")
                    if valor is None:
                        ui.notify("Valor inválido.", type="warning")
                        return
                    valores[chave] = valor
                saldo = SaldoInicial(titular=titular, ano=int(ano.value), **valores)
            try:
                await run.io_bound(saldos_iniciais.gravar, estado.banco, saldo, titular,
                                   estado.agora())  # fmt: skip
            except CoirError as erro:
                ui.notify(f"Não gravado: {erro}", type="negative")
                return
            estado.invalidar()
            ao_mudar()
            ui.notify("Saldos iniciais gravados." if saldo else "Saldos iniciais apagados.",
                      type="positive")  # fmt: skip

        with ui.row().classes("gap-2"):
            ui.button("Salvar", icon="save", on_click=salvar).mark("salvar-saldos")
            ui.button("Apagar", icon="delete_outline", on_click=lambda: salvar(apagar=True)).props(
                "flat")  # fmt: skip
    _custodia(estado, titular, ano, ao_mudar)


LINHAS_CUSTODIA = 5  # linhas em branco por vez


def _custodia(estado: Estado, titular: int, ano: ui.select, ao_mudar: Callable[[], object]) -> None:
    """Custódia em 31/12 (eventos "Custódia inicial"): os lançados e linhas para lançar vários
    ativos de uma vez, conferidos juntos pelo motor."""
    contas = eventos.contas_do_titular(estado.banco, titular)

    @ui.refreshable
    def lancados() -> None:
        lista = [e for e in eventos.listar(estado.banco, titular)
                 if e.dados.tipo is TipoEvento.CUSTODIA_INICIAL]  # fmt: skip
        if not lista:
            ui.label("Nenhuma custódia inicial lançada.").classes("coir-subtitulo text-sm")
        for e in lista:
            d = e.dados
            with ui.row().classes("items-center gap-3"):
                ui.label(f"{d.data:%d/%m/%Y} {e.conta_nome} {d.ticker}").classes(
                    "coir-numero w-[260px]")  # fmt: skip
                ui.label(f"{formatar(d.quantidade, T.INTEIRO)} ações · custo "
                         f"{_moeda(d.valor)}").classes("coir-numero")  # fmt: skip
                ui.button(icon="delete_outline", on_click=lambda _, i=e.id: excluir(i)).props(
                    "flat dense round").tooltip("Excluir")  # fmt: skip

    async def excluir(evento_id: int) -> None:
        try:
            await run.io_bound(eventos.excluir, estado.banco, titular, evento_id, estado.refs,
                               estado.fiscais)  # fmt: skip
        except CoirError as erro:
            ui.notify(f"Não excluído: {erro}", type="negative", multi_line=True)
            return
        estado.invalidar()
        ao_mudar()
        lancados.refresh()

    with ui.column().classes("coir-folha gap-2 w-full").mark("form-custodia"):
        ui.label("Custódia em 31/12").classes("coir-titulo text-base")
        ui.label(
            "Cada ativo que você tinha em 31/12 do ano acima, como no Bens e Direitos da "
            "declaração: conta, ticker, quantidade e custo total. Só ações, units, ETF, BDR e FII "
            "(opções, futuros e posição vendida pedem as notas)."
        ).classes("coir-subtitulo text-xs")
        if not contas:
            ui.label("Importe ao menos uma nota para criar a conta da corretora.").classes(
                "coir-vazio")  # fmt: skip
            return
        lancados()
        linhas = []
        for i in range(LINHAS_CUSTODIA):
            with ui.row().classes("items-center gap-2 no-wrap"):
                linhas.append((
                    ui.select(contas, value=next(iter(contas)), label="Conta").props(
                        "outlined dense").classes("w-[200px]"),
                    ui.input("Ticker").props("outlined dense").classes("w-[120px]").mark(
                        f"custodia-ticker-{i}"),
                    ui.input("Quantidade").props("outlined dense input-class=text-right"
                                                 ).classes("w-[130px]"),
                    ui.input("Custo total (R$)").props("outlined dense input-class=text-right"
                                                       ).classes("w-[170px]"),
                ))  # fmt: skip

        async def gravar() -> None:
            lote: list[tuple[eventos.DadosEvento, int | None]] = []
            dia = date(int(ano.value), 12, 31)
            for conta, ticker, quantidade, custo in linhas:
                if not (ticker.value or "").strip():
                    continue
                q, c = valor_br(quantidade.value or ""), valor_br(custo.value or "")
                if q is None or c is None:
                    ui.notify(f"{ticker.value}: quantidade ou custo inválido.", type="warning")
                    return
                lote.append((eventos.DadosEvento(
                    tipo=TipoEvento.CUSTODIA_INICIAL, data=dia, conta=int(conta.value),
                    ticker=ticker.value, cv="C", quantidade=q, valor=c), None))  # fmt: skip
            if not lote:
                ui.notify("Preencha ao menos um ticker.", type="warning")
                return
            try:
                await run.io_bound(eventos.salvar_lote, estado.banco, titular, lote, estado.refs,
                                   estado.fiscais)  # fmt: skip
            except CoirError as erro:
                ui.notify(f"Nada gravado: {erro}", type="negative", multi_line=True)
                return
            for _, ticker, quantidade, custo in linhas:
                for campo in (ticker, quantidade, custo):
                    campo.set_value("")
            estado.invalidar()
            ao_mudar()
            lancados.refresh()
            ui.notify(f"{len(lote)} ativo(s) lançado(s) na custódia de 31/12/{dia.year}.",
                      type="positive")  # fmt: skip

        ui.button("Gravar a custódia", icon="save", on_click=gravar).mark("gravar-custodia")


def _moeda(valor: Decimal) -> str:
    return formatar(valor, TipoColuna.DINHEIRO)


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
        t_saldos = ui.tab("Saldos iniciais").mark("aba-saldos")
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
            # nome próprio: ``valores`` é reaproveitado abaixo e o filtro leria a outra lista
            de_tickers = [(x.ticker, x.ativo, x.inicio, x.fim, x.isin) for x in tickers]
            busca = ui.input("Buscar ticker").props("outlined dense clearable").classes("w-[220px]")
            grade = tabela.grade(COLUNAS_TICKERS, de_tickers, altura="h-[560px]")
            grade.mark("grade-tickers")

            def filtrar(texto: str) -> None:
                vistos = [v for v in de_tickers if texto.strip().upper() in str(v[0])]
                tabela.atualizar(grade, COLUNAS_TICKERS, vistos)

            busca.on_value_change(lambda e: filtrar(e.value or ""))
            ui.label(
                "Trocas de ticker (data/trocas_ticker.csv) mantêm o mesmo ativo e o mesmo PM."
            ).classes("coir-subtitulo text-xs")
        with ui.tab_panel(t_saldos).classes("p-0 pt-3"):
            _saldos(estado, ao_mudar)
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
