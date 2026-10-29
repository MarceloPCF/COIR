"""Tela 8 — Opções: as séries por situação (abertas, vencidas/pó, exercidas, encerradas), com
posição, prêmio e resultado; ao escolher uma série, a trilha do ciclo com as notas."""

from collections.abc import Callable
from decimal import Decimal

from nicegui import run, ui

from coir.erros import CoirError
from coir.reports.modelo import Coluna, TipoColuna, Valor, formatar
from coir.servicos import carteira
from coir.servicos.derivativos import CicloOpcao
from coir.servicos.estado import Estado
from coir.ui import tabela

T = TipoColuna
ZERO = Decimal(0)
SITUACOES = (
    ("aberta", "Abertas"),
    ("vencida (pó)", "Vencidas (pó)"),
    ("exercida", "Exercidas"),
    ("encerrada", "Encerradas"),
)
COLUNAS = (
    Coluna("Série"),
    Coluna("Tipo"),
    Coluna("Objeto"),
    Coluna("Abertura", T.DATA),
    Coluna("Vencimento", T.DATA),
    Coluna("Posição", T.QUANTIDADE),
    Coluna("Prêmio", T.DINHEIRO),
    Coluna("Resultado", T.DINHEIRO),
)
COLUNAS_TRILHA = (
    Coluna("Data", T.DATA),
    Coluna("C/V"),
    Coluna("Qtd.", T.QUANTIDADE),
    Coluna("Valor", T.DINHEIRO),
    Coluna("Resultado", T.DINHEIRO),
    Coluna("Qtd. após", T.QUANTIDADE),
    Coluna("Regra"),
)


def _r(v: Decimal) -> str:
    return formatar(v, T.DINHEIRO)


def _lado(c: CicloOpcao) -> str:
    return ("Call" if c.call else "Put") + (" lançada" if c.posicao < 0 else " comprada")


def _valores(ciclos: list[CicloOpcao]) -> list[tuple[Valor, ...]]:
    return [
        (c.ticker, _lado(c), c.objeto, c.abertura, c.vencimento, c.posicao, c.premio,
         c.resultado if c.situacao != "aberta" or c.resultado else None)
        for c in ciclos
    ]  # fmt: skip


def do_ano(ciclos: list[CicloOpcao], ano: int | None) -> dict[str, list[CicloOpcao]]:
    """Abertas na data de referência e as demais encerradas no ano escolhido, por situação."""
    saida: dict[str, list[CicloOpcao]] = {s: [] for s, _ in SITUACOES}
    for c in ciclos:
        if c.situacao == "aberta" or c.fim.year == ano:
            saida[c.situacao].append(c)
    return saida


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Opções").classes("coir-titulo-pagina")
    ui.label(
        f"O ciclo de cada série: da abertura até zerar (venda, recompra, exercício ou vencimento "
        f"sem exercício, quando a opção vira pó). As abertas são as da data de referência; as "
        f"demais, as que terminaram em {estado.ano}. Prêmio = vendas - compras da série; no "
        "exercício o prêmio vai para o custo ou a venda do ativo-objeto (sem resultado na opção)."
    ).classes("coir-subtitulo")
    corpo = ui.column().classes("w-full gap-3")

    async def carregar() -> None:
        with corpo:
            espera = ui.spinner(size="lg")
        try:
            ciclos = await run.io_bound(estado.opcoes)
            mapa = await run.io_bound(estado.notas_das_operacoes)
            if corpo.is_deleted:
                return  # o usuário saiu da tela durante o cálculo
        except CoirError as erro:
            corpo.clear()
            with corpo:
                ui.label(f"Não foi possível calcular: {erro}").classes("coir-negativo")
            return
        espera.delete()
        if ciclos is None or mapa is None:
            return
        with corpo:
            grupos = do_ano(ciclos, estado.ano)
            if not any(grupos.values()):
                ui.label(f"Nenhuma opção aberta ou encerrada em {estado.ano}.").classes(
                    "coir-vazio w-full"
                )
                return
            _conteudo(grupos, mapa, estado.ano)

    ui.timer(0.01, carregar, once=True)


def _conteudo(grupos: dict[str, list[CicloOpcao]], mapa: dict[int, int], ano: int | None) -> None:
    encerrados = [c for s, lista in grupos.items() if s != "aberta" for c in lista]
    with ui.row().classes("w-full items-stretch gap-3").mark("cartoes-opcoes"):
        for rotulo, valor in (
            (f"Resultado de {ano}", sum((c.resultado for c in encerrados), ZERO)),
            ("Prêmio líquido das encerradas", sum((c.premio for c in encerrados), ZERO)),
            ("Prêmio das abertas", sum((c.premio for c in grupos["aberta"]), ZERO)),
        ):
            with ui.column().classes("coir-folha gap-0 min-w-[200px] flex-1"):
                ui.label(rotulo).classes("coir-campo-rotulo")
                ui.label(_r(valor)).classes("coir-numero text-2xl")
    inicial = next((s for s, _ in SITUACOES if grupos[s]), "aberta")
    with ui.tabs().props("no-caps align=left").classes("w-full") as abas:
        guias = {s: ui.tab(s, f"{r} ({len(grupos[s])})").mark(f"aba-{s.split()[0]}")
                 for s, r in SITUACOES}  # fmt: skip
    abas.value = guias[inicial]
    abas.on_value_change(lambda e: mostrar(_nome(e.value, inicial)))
    painel = ui.column().classes("w-full gap-2")
    trilha = ui.column().classes("w-full gap-2")

    def mostrar(situacao: str) -> None:
        lista = grupos[situacao]
        painel.clear()
        trilha.clear()
        with painel:
            if not lista:
                ui.label("Nenhuma série nesta situação.").classes("coir-vazio w-full")
                return
            total = ("", f"{len(lista)} séries", "", None, None, None,
                     sum((c.premio for c in lista), ZERO),
                     sum((c.resultado for c in lista), ZERO))  # fmt: skip
            tabela.grade(
                COLUNAS, _valores(lista), chaves=list(range(len(lista))), total=total,
                altura="h-[420px]", ao_clicar=_abrir(lista, mapa, trilha),
            ).mark("grade-opcoes")  # fmt: skip
        with trilha:
            ui.label("Escolha uma série para ver a trilha do ciclo.").classes("coir-vazio w-full")

    mostrar(inicial)


def _nome(valor: object, padrao: str) -> str:
    """A aba escolhida volta como nome (cliente) ou como o próprio elemento."""
    if isinstance(valor, ui.tab):
        return str(valor.props["name"])
    return str(valor) if valor else padrao


def _abrir(
    lista: list[CicloOpcao], mapa: dict[int, int], destino: ui.column
) -> Callable[[object], None]:
    def abrir(chave: object) -> None:
        c = lista[int(str(chave))]
        destino.clear()
        with destino:
            _trilha(c, mapa)

    return abrir


def _trilha(c: CicloOpcao, mapa: dict[int, int]) -> None:
    with ui.row().classes("items-baseline gap-3"):
        ui.label(f"Trilha de {c.ticker}").classes("coir-titulo text-lg")
        ui.label(f"{_lado(c)} · objeto {c.objeto} · {c.situacao}").classes("coir-subtitulo text-sm")
    valores: list[tuple[Valor, ...]] = [
        (m.data, m.cv or "", m.quantidade, m.valor, m.resultado or None, m.quantidade_apos,
         m.regra)
        for m in c.passos
    ]  # fmt: skip
    notas = [carteira.nota_de(m.origem, mapa) for m in c.passos]
    tabela.grade(
        COLUNAS_TRILHA,
        valores,
        chaves=notas,
        pelo_conteudo=True,
        ao_clicar=lambda n: n is not None and ui.navigate.to(f"/notas/{n}"),
    ).mark("trilha-opcao")
    ui.label("Clique num passo para abrir a nota de origem.").classes("coir-subtitulo text-xs")
