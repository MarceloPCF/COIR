"""Janela do COIR: barra lateral, topo com titular e ano, e as telas como sub-páginas.

A interface só apresenta: os dados vêm de `coir.servicos` (que chamam o motor) e os relatórios
do modelo neutro de `coir.reports`.
"""

from collections.abc import Callable
from dataclasses import dataclass

from nicegui import ui
from nicegui.events import ValueChangeEventArguments

from coir.servicos.estado import Estado
from coir.ui import preferencias, tema
from coir.ui.telas import (
    apuracao,
    assistente,
    cadastros,
    carteira,
    configuracoes,
    darf,
    importar,
    notas,
    operacoes,
    painel,
    relatorios,
)

Avisar = Callable[[], object]  # a tela mudou o banco: o topo (titulares, anos) se refaz
Montar = Callable[[Estado, Avisar], None]


@dataclass(frozen=True)
class Tela:
    rota: str
    rotulo: str
    icone: str
    montar: Montar
    exige_dados: bool = True  # sem banco ou titular, abre o assistente no lugar


GRUPOS: tuple[tuple[str, tuple[Tela, ...]], ...] = (
    ("", (Tela("/", "Painel", "space_dashboard", painel.montar),)),
    ("Notas", (
        Tela("/importar", "Importar", "upload_file", importar.montar),
        Tela("/notas", "Notas", "receipt_long", notas.montar),
        Tela("/operacoes", "Operações", "list_alt", operacoes.montar),
    )),
    ("Imposto", (
        Tela("/apuracao", "Apuração mensal", "calculate", apuracao.montar),
        Tela("/darf", "DARF", "request_quote", darf.montar),
        Tela("/relatorios", "Relatórios", "description", relatorios.montar),
    )),
    ("Carteira", (
        Tela("/carteira", "Carteira e preço médio", "account_balance_wallet", carteira.montar),
    )),
    ("Sistema", (
        Tela("/cadastros", "Cadastros", "badge", cadastros.montar, exige_dados=False),
        Tela("/configuracoes", "Configurações", "settings", configuracoes.montar,
             exige_dados=False),
    )),
)  # fmt: skip
TELAS = tuple(t for _, telas in GRUPOS for t in telas)
INICIO = Tela("/inicio", "Primeiros passos", "flag", assistente.montar, exige_dados=False)


def _ir(rota: str) -> Callable[[], None]:
    return lambda: ui.navigate.to(rota)


def _construtor(tela: Tela, estado: Estado, avisar: Avisar) -> Callable[[], None]:
    def construir() -> None:
        if tela.exige_dados and not estado.pronto:
            assistente.montar(estado, avisar)
        else:
            tela.montar(estado, avisar)

    return construir


def _nota(estado: Estado, avisar: Avisar) -> Callable[[str], None]:
    """Uma nota direto pelo endereço (/notas/<id>), vinda da tela de operações."""

    def construir(nota: str) -> None:
        if not estado.pronto:
            assistente.montar(estado, avisar)
        elif nota.isdigit():
            notas.montar(estado, avisar, int(nota))
        else:
            ui.label(f"Nota {nota!r} inválida.").classes("coir-vazio w-full")

    return construir


def montar(estado: Estado) -> None:
    """Página raiz (uma por janela)."""
    tema.aplicar()
    preferencias.aplicar_tema()  # sistema, claro ou escuro (Configurações)
    estado.escolher_padrao(estado.agora().date())
    paginas: ui.sub_pages | None = None

    def recarregar() -> None:
        if paginas is not None:
            paginas.refresh()

    def escolher(atributo: str) -> Callable[[ValueChangeEventArguments], None]:
        def mudar(e: ValueChangeEventArguments) -> None:
            setattr(estado, atributo, e.value)
            recarregar()

        return mudar

    @ui.refreshable
    def seletores() -> None:
        hoje = estado.agora().date()
        estado.escolher_padrao(hoje)
        titulares = {t.id: t.nome for t in estado.titulares()}
        if titulares:
            ui.select(titulares, label="Titular", value=estado.titular,
                      on_change=escolher("titular"),
                      ).props("outlined dense options-dense").classes("min-w-[220px]")  # fmt: skip
        ui.select(estado.anos(hoje), label="Ano", value=estado.ano, on_change=escolher("ano"),
                  ).props("outlined dense options-dense").classes("w-[110px]")  # fmt: skip

    with ui.header().classes("coir-topo items-center gap-4 px-4 py-2"):
        ui.label("COIR").classes("coir-marca")
        ui.space()
        with ui.row().classes("items-center gap-4"):
            seletores()
        ui.button(icon="contrast", on_click=preferencias.alternar_tema).props(
            "flat round").tooltip("Tema claro ou escuro")  # fmt: skip

    with (
        ui.left_drawer(value=True, fixed=True, bordered=False)
        .props("width=232 breakpoint=700 show-if-above")
        .classes("coir-lateral p-0")
    ):
        for grupo, telas in GRUPOS:
            if grupo:
                ui.label(grupo).classes("coir-grupo")
            for t in telas:
                with ui.item(on_click=_ir(t.rota)).props("clickable").mark(f"menu{t.rota}"):
                    with ui.item_section().props("avatar").classes("min-w-0 pr-3"):
                        ui.icon(t.icone, size="20px")
                    ui.item_section(t.rotulo)

    with ui.column().classes("w-full max-w-[1400px] mx-auto px-6 py-5 gap-4"):
        rotas: dict[str, Callable[..., None]] = {
            t.rota: _construtor(t, estado, seletores.refresh) for t in (*TELAS, INICIO)
        }
        rotas["/notas/{nota}"] = _nota(estado, seletores.refresh)
        paginas = ui.sub_pages(rotas).classes("w-full")
