"""Tela 14 — Declaração anual (IRPF): as fichas na ordem do programa da Receita, cada valor com
botão de copiar e cada item com a marcação "lançado no programa" (gravada por titular e ano), e a
conferência com a declaração entregue quando o PDF está em <dados>/_dirpf (escolha do usuário de
20/10/2026: alternativa COIR, ficha a ficha)."""

from collections.abc import Awaitable, Callable
from decimal import Decimal

from nicegui import run, ui

from coir.domain.cpf import formatar_cnpj
from coir.domain.irpf import FichasIrpf, LinhaFicha, MapeamentoIrpf, MesFichaIrpf, RendimentoIrpf
from coir.erros import CoirError
from coir.reports.irpf import ALIQUOTA, MESES, aliquota
from coir.reports.modelo import Coluna, TipoColuna, formatar
from coir.servicos.declaracao import Conferencia
from coir.servicos.estado import Estado
from coir.ui import tabela

T = TipoColuna
FICHAS = (
    ("bens", "Bens e Direitos"),
    ("dividas", "Dívidas e ônus reais"),
    ("isentos", "Rendimentos isentos"),
    ("exclusivos", "Tributação exclusiva"),
    ("rv", "Renda variável"),
    ("fii", "FII e Fiagro"),
    ("pago", "Imposto pago/retido"),
    ("conferencia", "Conferir com a DIRPF"),
)


def _programa(valor: Decimal) -> str:
    """Valor como se digita no programa: sem separador de milhar ("1234,56")."""
    return formatar(valor, T.DINHEIRO, milhar=False)


def _copiar(texto: str, rotulo: str = "") -> None:
    def copiar() -> None:
        ui.clipboard.write(texto)
        ui.notify(f"Copiado: {texto[:60]}", type="positive", timeout=1500)

    with ui.row().classes("items-center gap-1 no-wrap"):
        if rotulo:
            ui.label(rotulo).classes("coir-numero")
        ui.button(icon="content_copy", on_click=copiar).props("flat dense round size=sm").tooltip(
            "Copiar para colar no programa IRPF")  # fmt: skip


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Declaração anual (IRPF)").classes("coir-titulo-pagina")
    ui.label(
        f"As fichas de renda variável do ano-calendário {estado.ano} na ordem do programa da "
        "Receita: copie cada valor e marque o que já lançou. Os valores vêm do motor (custo médio "
        "do titular, apuração mensal, proventos)."
    ).classes("coir-subtitulo")
    corpo = ui.column().classes("w-full gap-3")

    async def carregar() -> None:
        corpo.clear()
        with corpo:
            espera = ui.spinner(size="lg")
        try:
            resultado = await run.io_bound(estado.fichas_irpf)
            feitos = await run.io_bound(estado.lancados)
        except CoirError as erro:
            corpo.clear()
            with corpo:
                ui.label(f"Não foi possível montar a declaração: {erro}").classes("coir-negativo")
            return
        if resultado is None or feitos is None:
            return  # a página foi fechada durante o cálculo
        f, mapa = resultado
        espera.delete()
        with corpo:
            if f is None or mapa is None:
                exercicio = (estado.ano or 0) + 1
                ui.label(
                    f"Sem o mapeamento das fichas do programa IRPF {exercicio} "
                    f"(data/irpf/{exercicio}.yaml)."
                ).classes("coir-vazio w-full")
                return
            _tela(estado, f, mapa, feitos)

    ui.timer(0.01, carregar, once=True)


def _tela(estado: Estado, f: FichasIrpf, mapa: MapeamentoIrpf, feitos: set[str]) -> None:
    provisorio = " · mapeamento provisório" if f.provisorio else ""
    ui.label(f"Exercício {f.exercicio}, ano-calendário {f.ano}{provisorio}").classes(
        "coir-titulo text-lg"
    )
    for aviso in f.avisos:
        ui.label(aviso).classes("coir-carimbo text-xs")

    async def marcar(chave: str, valor: bool) -> None:
        try:
            await run.io_bound(estado.marcar_lancado, chave, valor)
        except CoirError as erro:
            ui.notify(f"Não gravado: {erro}", type="negative")
            return
        if valor:
            feitos.add(chave)
        else:
            feitos.discard(chave)

    def lancado(chave: str) -> None:
        ui.checkbox(
            "lançado", value=chave in feitos, on_change=lambda e: marcar(chave, bool(e.value))
        ).mark(f"lancado-{chave}")

    with ui.row().classes("w-full items-start gap-4 no-wrap"):
        with ui.column().classes("w-[220px] shrink-0 gap-0 coir-folha p-1"):
            itens = {chave: ui.item(rotulo).props("clickable dense").classes("rounded")
                     .mark(f"ficha-{chave}") for chave, rotulo in FICHAS}  # fmt: skip
        painel = ui.column().classes("grow min-w-0 gap-2")

    def mostrar(chave: str) -> Callable[[], Awaitable[None] | None]:
        async def abrir() -> None:
            for c, item in itens.items():
                item.classes(add="coir-item-ativo" if c == chave else "",
                             remove="" if c == chave else "coir-item-ativo")  # fmt: skip
            painel.clear()
            with painel:
                if chave == "conferencia":
                    await _conferencia(estado)
                else:
                    _ficha(chave, f, mapa, lancado)

        return abrir

    for chave, item in itens.items():
        item.on_click(mostrar(chave))
    ui.timer(0.01, mostrar("bens"), once=True)


def _ficha(chave: str, f: FichasIrpf, mapa: MapeamentoIrpf,
           lancado: Callable[[str], None]) -> None:  # fmt: skip
    if chave == "bens":
        if not f.bens:
            ui.label("Nenhum bem de renda variável no ano.").classes("coir-vazio w-full")
        for b in f.bens:
            item = b.ticker or f"{b.grupo}.{b.codigo} {b.cnpj or b.discriminacao[:40]}"
            with ui.column().classes("coir-folha w-full gap-1").mark(f"bem-{item}"):
                with ui.row().classes("items-center gap-3 w-full"):
                    ui.label(f"{b.grupo}.{b.codigo}").classes("coir-titulo")
                    if b.ticker:
                        _copiar(b.ticker, f"código {b.ticker}")
                    if b.cnpj:
                        _copiar(formatar_cnpj(b.cnpj), f"CNPJ {formatar_cnpj(b.cnpj)}")
                    ui.space()
                    lancado(f"bens:{item}")
                with ui.row().classes("items-center gap-6"):
                    _copiar(_programa(b.anterior), f"31/12/{f.ano - 1}: {_programa(b.anterior)}")
                    _copiar(_programa(b.atual), f"31/12/{f.ano}: {_programa(b.atual)}")
                with ui.row().classes("items-start gap-2 no-wrap w-full"):
                    ui.label(b.discriminacao).classes("coir-subtitulo text-sm grow")
                    _copiar(b.discriminacao)
    elif chave == "dividas":
        if not f.dividas:
            ui.label("Nenhuma posição vendida em 31/12.").classes("coir-vazio w-full")
        for x in f.dividas:
            with ui.column().classes("coir-folha w-full gap-1"):
                with ui.row().classes("items-center gap-3 w-full"):
                    ui.label(f"{x.codigo} · {x.ticker}").classes("coir-titulo")
                    ui.space()
                    lancado(f"dividas:{x.ticker}")
                with ui.row().classes("items-center gap-6"):
                    _copiar(_programa(x.anterior), f"31/12/{f.ano - 1}: {_programa(x.anterior)}")
                    _copiar(_programa(x.atual), f"31/12/{f.ano}: {_programa(x.atual)}")
                with ui.row().classes("items-start gap-2 no-wrap w-full"):
                    ui.label(x.discriminacao).classes("coir-subtitulo text-sm grow")
                    _copiar(x.discriminacao)
    elif chave in ("isentos", "exclusivos", "pago"):
        ficha = "imposto_pago" if chave == "pago" else chave
        _rendimentos([r for r in f.rendimentos if r.ficha == ficha], chave, mapa, lancado)
    elif chave == "rv":
        _meses(f.renda_variavel, "rv", lancado, duas=True)
    elif chave == "fii":
        if not f.fii:
            ui.label("Sem operações com FII ou Fiagro no ano.").classes("coir-vazio w-full")
        _meses(f.fii, "fii", lancado, duas=False)


def _rendimentos(rs: list[RendimentoIrpf], chave: str, mapa: MapeamentoIrpf,
                 lancado: Callable[[str], None]) -> None:  # fmt: skip
    if not rs:
        ui.label("Nada nesta ficha no ano.").classes("coir-vazio w-full")
    secao = {"isentos": mapa.isentos, "exclusivos": mapa.exclusivos,
             "pago": mapa.imposto_pago}[chave]  # fmt: skip
    for r in rs:
        descricao = secao.linha(r.chave).descricao
        with ui.row().classes("coir-folha w-full items-center gap-3"):
            ui.label(r.codigo).classes("coir-titulo w-[36px]")
            with ui.column().classes("grow gap-0"):
                ui.label(r.fonte).classes("text-sm")
                ui.label(descricao).classes("coir-subtitulo text-xs")
            if r.cnpj:
                _copiar(formatar_cnpj(r.cnpj), formatar_cnpj(r.cnpj))
            _copiar(_programa(r.valor), _programa(r.valor))
            lancado(f"{chave}:{r.codigo}:{r.cnpj or r.chave}")


def _meses(meses: tuple[MesFichaIrpf, ...], chave: str, lancado: Callable[[str], None],
           duas: bool) -> None:  # fmt: skip
    """Um mês por aba; só as linhas que se digitam no programa têm botão de copiar (o resto o
    programa calcula: serve para conferir)."""
    if not meses:
        return
    with ui.tabs().classes("w-full") as abas:
        tabs = {m.mes: ui.tab(MESES[m.mes - 1][:3]) for m in meses}
    with ui.tab_panels(abas, value=tabs[meses[0].mes]).classes("w-full"):
        for m in meses:
            with ui.tab_panel(tabs[m.mes]):
                with ui.row().classes("w-full items-center"):
                    ui.label(MESES[m.mes - 1]).classes("coir-titulo")
                    ui.space()
                    lancado(f"{chave}:{m.mes:02d}")
                for x in m.linhas:
                    _linha_mes(x, x.digita, duas)


def _linha_mes(x: LinhaFicha, digita: bool, duas: bool) -> None:
    def texto(v: Decimal) -> str:
        return aliquota(v) if x.rotulo == ALIQUOTA else _programa(v)

    with ui.row().classes("w-full items-center gap-3 no-wrap"):
        ui.label(x.rotulo).classes("grow text-sm" if digita else "grow text-sm coir-subtitulo")
        for v in (x.comuns, *((x.day_trade,) if duas and x.day_trade is not None else ())):
            if digita and v:
                _copiar(texto(v), texto(v))
            else:
                ui.label(texto(v)).classes("coir-numero w-[120px] text-right coir-subtitulo")
        if duas and x.day_trade is None:
            ui.label("").classes("w-[120px]")


COLUNAS_CONFERENCIA = (
    Coluna("Ficha"), Coluna("Item"), Coluna("Coluna"), Coluna("COIR"), Coluna("DIRPF"),
    Coluna("Diferença", T.DINHEIRO), Coluna("Situação"), Coluna("Explicação"),
)  # fmt: skip


async def _conferencia(estado: Estado) -> None:
    ui.label(
        'Compara as fichas com a declaração entregue (PDF "Imprimir > Declaração" do programa '
        "IRPF, salvo em _dirpf na pasta de dados). As divergências já analisadas ficam em "
        "gabarito/dirpf-explicadas.csv."
    ).classes("coir-subtitulo text-xs")
    espera = ui.spinner()
    try:
        conferencia: Conferencia | None = await run.io_bound(estado.conferir_dirpf)
    except CoirError as erro:
        espera.delete()
        ui.label(f"Não foi possível ler a declaração: {erro}").classes("coir-negativo")
        return
    espera.delete()
    if conferencia is None:
        ui.label("Nenhuma DIRPF deste ano-calendário em _dirpf.").classes("coir-vazio w-full")
        return
    total = len(conferencia.linhas)
    abertas = conferencia.abertas
    ui.label(f"{conferencia.arquivo}: {total - len(abertas)} de {total} itens conferem ou estão "
             f"explicados; {len(abertas)} divergências em aberto.").classes(
        "coir-titulo text-base").mark("resumo-conferencia")  # fmt: skip
    valores = [
        (
            x.secao,
            x.item,
            x.coluna,
            str(x.coir if x.coir is not None else ""),
            str(x.dirpf if x.dirpf is not None else ""),
            x.diferenca,
            x.classe,
            x.explicacao,
        )
        for x in conferencia.linhas
        if x.classe != "confere"
    ]
    tabela.grade(COLUNAS_CONFERENCIA, valores, altura="h-[520px]").mark("grade-conferencia")
