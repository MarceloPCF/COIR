"""Tela 1 — Assistente inicial, em três passos: pasta de dados, titular e primeira importação.

Abre sozinho enquanto não houver banco ou titular. A custódia inicial e o prejuízo de anos
anteriores ficam para os lançamentos manuais (F7).
"""

from collections.abc import Callable
from pathlib import Path

from nicegui import ui

from coir.domain.cpf import cpf_valido
from coir.erros import CoirError
from coir.servicos import cadastros
from coir.servicos.estado import Estado
from coir.ui.telas import importar


def _contar_pdfs(pasta: Path) -> str:
    notas = pasta / "_notas-corretagens"
    if not pasta.is_dir():
        return "A pasta ainda não existe: o COIR cria ao continuar."
    qtd = sum(1 for p in notas.glob("*") if p.suffix.lower() == ".pdf") if notas.is_dir() else 0
    banco = "com banco do COIR" if (pasta / "coir.db").is_file() else "sem banco ainda"
    return f"Pasta encontrada, {banco}; {qtd} PDFs em _notas-corretagens."


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Bem-vindo ao COIR").classes("coir-titulo-pagina")
    ui.label(
        "Três passos para começar. Tudo fica nesta pasta, neste computador: nada da sua "
        "carteira sai daqui."
    ).classes("coir-subtitulo")

    with ui.stepper().props("flat animated").classes("coir-folha w-full max-w-[880px]") as passos:
        with ui.step("Pasta de dados", icon="folder"):
            ui.label("Onde ficam os seus dados?").classes("coir-titulo text-base")
            ui.label("As notas em PDF, o banco do COIR, os backups e os relatórios.").classes(
                "coir-subtitulo"
            )
            campo = (
                ui.input("Pasta de dados", value=str(estado.dados))
                .props("outlined dense")
                .classes("w-full")
                .mark("pasta-dados")
            )
            situacao = ui.label(_contar_pdfs(estado.dados)).classes("coir-subtitulo text-sm")
            campo.on_value_change(lambda e: situacao.set_text(_contar_pdfs(Path(e.value or "."))))

            def confirmar_pasta() -> None:
                texto = (campo.value or "").strip()
                if not texto:
                    ui.notify("Informe a pasta de dados.", type="warning")
                    return
                try:
                    estado.trocar_dados(Path(texto).expanduser())
                    estado.preparar(estado.agora())
                except (CoirError, OSError) as erro:
                    ui.notify(f"Não foi possível usar a pasta: {erro}", type="negative")
                    return
                ao_mudar()
                mostrar_titular.refresh()
                area_notas.refresh()
                passos.next()

            with ui.stepper_navigation():
                ui.button("Continuar", on_click=confirmar_pasta).mark("continuar-pasta")

        with ui.step("Titular", icon="person"):

            @ui.refreshable
            def mostrar_titular() -> None:
                titulares = estado.titulares()
                if titulares:
                    ui.label(
                        "Titular já cadastrado: " + ", ".join(t.nome for t in titulares)
                    ).classes("text-base")
                    ui.label(
                        "Outros titulares entram em Cadastros (ou sozinhos, pelo CPF impresso "
                        "na nota)."
                    ).classes("coir-subtitulo text-sm")
                    with ui.stepper_navigation():
                        ui.button("Voltar", on_click=passos.previous).props("flat")
                        ui.button("Continuar", on_click=passos.next).mark("continuar-titular")
                    return
                ui.label("Quem é o titular das notas?").classes("coir-titulo text-base")
                ui.label(
                    "O CPF separa as contas de cada titular; ele nunca vai para nomes de "
                    "arquivo nem para o registro de atividades."
                ).classes("coir-subtitulo text-sm")
                nome = ui.input("Nome").props("outlined dense").classes("w-full").mark("nome")
                cpf = (
                    ui.input("CPF", validation={"CPF inválido": lambda v: cpf_valido(v or "")})
                    .props("outlined dense mask='###.###.###-##' lazy-rules")
                    .classes("w-[220px]")
                    .mark("cpf")
                )

                def cadastrar() -> None:
                    try:
                        cadastros.adicionar_titular(estado.banco, cpf.value or "", nome.value or "")
                    except CoirError as erro:
                        ui.notify(str(erro), type="negative")
                        return
                    estado.invalidar()
                    estado.escolher_padrao(estado.agora().date())
                    ao_mudar()
                    mostrar_titular.refresh()
                    passos.next()

                with ui.stepper_navigation():
                    ui.button("Voltar", on_click=passos.previous).props("flat")
                    ui.button("Cadastrar e continuar", on_click=cadastrar).mark("cadastrar")

            mostrar_titular()

        with ui.step("Notas", icon="upload_file"):
            ui.label("Importe as notas de corretagem").classes("coir-titulo text-base")
            ui.label("Pode ser agora ou depois, pela tela Importar notas.").classes(
                "coir-subtitulo text-sm"
            )

            @ui.refreshable
            def area_notas() -> None:  # a pasta das notas muda no passo 1
                importar.area(estado, ao_mudar)

            area_notas()
            with ui.stepper_navigation():
                ui.button("Voltar", on_click=passos.previous).props("flat")
                ui.button("Concluir", on_click=lambda: ui.navigate.to("/")).mark("concluir")
