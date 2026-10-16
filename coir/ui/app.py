"""Janela do COIR: barra lateral, topo com titular e ano, e as telas como sub-páginas.

A interface só apresenta: os dados vêm de `coir.servicos` (que chamam o motor) e os relatórios
do modelo neutro de `coir.reports`.
"""

from collections.abc import Callable
from dataclasses import dataclass
from datetime import date

from nicegui import ui

from coir.servicos.estado import Estado
from coir.ui import tema
from coir.ui.telas import provisoria


@dataclass(frozen=True)
class Tela:
    rota: str
    rotulo: str
    icone: str
    montar: Callable[[Estado], None]
    descricao: str = ""


def _provisoria(titulo: str, descricao: str) -> Callable[[Estado], None]:
    return lambda estado: provisoria.montar(titulo, descricao)


GRUPOS: tuple[tuple[str, tuple[Tela, ...]], ...] = (
    ("", (Tela("/", "Painel", "space_dashboard", _provisoria(
        "Painel", "Resumo do ano do titular: DARF do mês, avisos e evolução do resultado.")),)),
    ("Notas", (
        Tela("/importar", "Importar notas", "upload_file", _provisoria(
            "Importar notas", "Escolha os PDFs das notas de corretagem; o COIR confere cada nota "
            "antes de gravar.")),
        Tela("/notas", "Notas", "receipt_long", _provisoria(
            "Notas", "As notas gravadas, com os dados extraídos ao lado da página do PDF.")),
        Tela("/operacoes", "Operações", "list_alt", _provisoria(
            "Operações", "O livro de operações do ano, com filtros por ativo, corretora e "
            "mercado.")),
    )),
    ("Imposto", (
        Tela("/apuracao", "Apuração mensal", "calculate", _provisoria(
            "Apuração mensal", "O resultado de cada mês por categoria e a memória de cálculo.")),
        Tela("/darf", "DARF", "request_quote", _provisoria(
            "DARF", "As DARFs do ano, o vencimento, o pagamento e a multa por atraso.")),
        Tela("/relatorios", "Relatórios", "description", _provisoria(
            "Relatórios", "Os relatórios do ano na tela e a exportação em xlsx, csv e pdf.")),
    )),
    ("Carteira", (
        Tela("/carteira", "Carteira e preço médio", "account_balance_wallet", _provisoria(
            "Carteira e preço médio", "A posição numa data e a trilha do preço médio de cada "
            "ativo.")),
    )),
    ("Sistema", (
        Tela("/cadastros", "Cadastros", "badge", _provisoria(
            "Cadastros", "Titulares, contas, corretoras, ativos e parâmetros fiscais.")),
        Tela("/configuracoes", "Configurações", "settings", _provisoria(
            "Configurações", "Pasta de dados, tema, backups e o registro de atividades.")),
    )),
)  # fmt: skip
TELAS = tuple(t for _, telas in GRUPOS for t in telas)


def _ir(rota: str) -> Callable[[], None]:
    return lambda: ui.navigate.to(rota)


def _construtor(tela: Tela, estado: Estado) -> Callable[[], None]:
    return lambda: tela.montar(estado)


def montar(estado: Estado, hoje: Callable[[], date] = date.today) -> None:
    """Página raiz (uma por janela)."""
    tema.aplicar()
    escuro = ui.dark_mode(value=None)  # segue o sistema até o usuário escolher
    estado.escolher_padrao(hoje())
    paginas: ui.sub_pages | None = None

    def recarregar() -> None:
        if paginas is not None:
            paginas.refresh()

    with ui.header().classes("coir-topo items-center gap-4 px-4 py-2"):
        ui.label("COIR").classes("coir-marca")
        ui.space()
        titulares = {t.id: t.nome for t in estado.titulares()}
        if titulares:
            ui.select(titulares, label="Titular", value=estado.titular,
                      on_change=lambda e: (setattr(estado, "titular", e.value), recarregar()),
                      ).props("outlined dense options-dense").classes("min-w-[220px]")  # fmt: skip
        ui.select(estado.anos(hoje()), label="Ano", value=estado.ano,
                  on_change=lambda e: (setattr(estado, "ano", e.value), recarregar()),
                  ).props("outlined dense options-dense").classes("w-[110px]")  # fmt: skip
        ui.button(icon="contrast", on_click=lambda: escuro.set_value(not escuro.value)).props(
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
        paginas = ui.sub_pages({t.rota: _construtor(t, estado) for t in TELAS})
