"""Tela 18 — Configurações, em seções: pasta de dados, aparência, backups (fazer agora e
restaurar, com cópia do banco atual antes) e o registro de atividades (sem CPF nem nome)."""

from collections.abc import Callable

from nicegui import run, ui

from coir.config import ARQUIVO_REGISTRO, ler_registro
from coir.erros import CoirError
from coir.servicos import manutencao
from coir.servicos.estado import Estado
from coir.ui import preferencias


def _tamanho(n: int) -> str:
    return f"{n / 1_048_576:.1f} MB".replace(".", ",") if n >= 1_048_576 else f"{n // 1024} KB"


def _secao(titulo: str, texto: str) -> ui.column:
    with ui.column().classes("coir-folha w-full gap-2") as coluna:
        ui.label(titulo).classes("coir-titulo text-base")
        ui.label(texto).classes("coir-subtitulo text-sm")
    return coluna


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Configurações").classes("coir-titulo-pagina")
    ui.label("Onde ficam os dados, a aparência, os backups e o registro de atividades.").classes(
        "coir-subtitulo"
    )
    with (
        _secao("Pasta de dados", "Notas em PDF, banco, backups, comprovantes e relatórios."),
        ui.row().classes("items-center gap-3"),
    ):
        ui.label(str(estado.dados)).classes("coir-numero").mark("pasta-atual")
        ui.button(
            "Trocar a pasta", icon="folder", on_click=lambda: ui.navigate.to("/inicio")
        ).props("flat")

    with _secao("Aparência", "O tema segue o Windows ou fica fixo; vale nas próximas aberturas."):
        ui.radio(
            {"sistema": "Como o sistema", "claro": "Claro", "escuro": "Escuro"},
            value=preferencias.tema(),
            on_change=lambda e: preferencias.escolher_tema(e.value),
        ).props("inline").mark("tema")

    with _secao(
        "Backups",
        "O COIR copia o banco antes de importar e de migrar (os 10 mais recentes ficam). "
        "Restaurar volta o banco àquele momento; o banco atual vira um backup antes, para "
        "poder desfazer.",
    ):
        ui.button("Fazer backup agora", icon="backup", on_click=lambda: fazer()).mark(
            "fazer-backup"
        )
        lista = ui.column().classes("w-full gap-0")

    def mostrar_backups() -> None:
        lista.clear()
        with lista:
            copias = manutencao.backups(estado.banco)
            if not copias:
                ui.label("Nenhum backup ainda.").classes("coir-subtitulo text-sm")
            for b in copias:
                with ui.row().classes("coir-item-lista w-full items-center gap-4 py-1"):
                    quando = f"{b.feito_em:%d/%m/%Y %H:%M:%S}" if b.feito_em else b.caminho.name
                    ui.label(quando).classes("coir-numero w-[170px]")
                    ui.label(_tamanho(b.tamanho)).classes("coir-subtitulo text-sm w-[80px]")
                    ui.space()
                    ui.button("Restaurar", icon="restore", on_click=_confirmar(b, restaurar)).props(
                        "flat dense"
                    ).mark(f"restaurar-{b.caminho.name}")

    async def fazer() -> None:
        try:
            destino = await run.io_bound(manutencao.fazer_agora, estado.banco, estado.agora())
        except CoirError as erro:
            ui.notify(str(erro), type="negative")
            return
        if destino is not None:
            ui.notify(f"Backup feito: {destino.name}", type="positive")
        mostrar_backups()

    async def restaurar(b: manutencao.Backup) -> None:
        try:
            await run.io_bound(manutencao.restaurar, estado.banco, b.caminho, estado.agora())
        except CoirError as erro:
            ui.notify(f"Restauração não feita: {erro}", type="negative", multi_line=True)
            return
        estado.invalidar()
        estado.escolher_padrao(estado.agora().date())
        ao_mudar()
        ui.notify("Banco restaurado. O anterior ficou na lista de backups.", type="positive")
        mostrar_backups()

    mostrar_backups()

    with _secao(
        "Registro de atividades",
        f"As últimas linhas de {ARQUIVO_REGISTRO.as_posix()} na pasta de dados: importações, "
        "backups, migrações e erros. O registro mascara CPF e não grava os dados das consultas.",
    ):
        registro = ui.log(max_lines=300).classes("w-full h-[320px] text-xs").mark("registro")

        def ler() -> None:
            registro.clear()
            for linha in ler_registro(estado.dados) or ["(ainda sem registro)"]:
                registro.push(linha)

        ui.button("Atualizar", icon="refresh", on_click=ler).props("flat dense")
        ler()


def _confirmar(
    b: manutencao.Backup, restaurar: Callable[[manutencao.Backup], object]
) -> Callable[[], None]:
    def abrir() -> None:
        quando = f"{b.feito_em:%d/%m/%Y %H:%M:%S}" if b.feito_em else b.caminho.name
        with ui.dialog() as janela, ui.card().classes("max-w-[480px]"):
            ui.label("Restaurar o backup?").classes("coir-titulo text-base")
            ui.label(
                f"O banco volta a {quando}. O que foi importado ou lançado depois some do "
                "banco (o banco atual vira um backup antes, para desfazer).",
            ).classes("text-sm")

            async def sim() -> None:
                janela.close()
                resultado = restaurar(b)
                if hasattr(resultado, "__await__"):
                    await resultado  # type: ignore[misc]

            with ui.row().classes("w-full justify-end"):
                ui.button("Cancelar", on_click=janela.close).props("flat")
                ui.button("Restaurar", on_click=sim).props("color=negative").mark(
                    "confirmar-restaurar"
                )
        janela.open()

    return abrir
