"""Escolha da pasta das notas e da pasta das planilhas da B3 (tela Importar)."""

from collections.abc import Callable
from pathlib import Path

from nicegui import ui

from coir.erros import ConfiguracaoError
from coir.servicos.estado import Estado
from coir.ui.dialogos import escolher_pasta


def seletor(estado: Estado, chave: str, rotulo: str, atual: Path,
            recarregar: Callable[[], object]) -> None:  # fmt: skip
    """Mostra a pasta em uso e permite trocar por outra (ex.: a pasta onde você já guarda as
    notas) ou voltar à padrão, dentro da pasta de dados."""
    with ui.row().classes("items-center gap-2"):
        ui.label(f"{rotulo}: {atual}").classes("coir-subtitulo text-xs").mark(f"{chave}-atual")
        with ui.expansion("Trocar", icon="drive_folder_upload").props("dense").classes(
            "text-xs"
        ):  # fmt: skip
            campo = ui.input("Caminho da pasta", value=str(atual)).props(
                "outlined dense").classes("w-[460px]").mark(f"{chave}-campo")  # fmt: skip

            def usar(pasta: Path | None) -> None:
                try:
                    estado.escolher_pasta(chave, pasta)
                except ConfiguracaoError as erro:
                    ui.notify(str(erro), type="warning")
                    return
                ui.notify("Pasta trocada." if pasta else "Voltou à pasta padrão.", type="positive")
                recarregar()

            async def procurar() -> None:
                escolhida = await escolher_pasta(Path(campo.value or str(atual)))
                if escolhida is not None:
                    campo.set_value(str(escolhida))
                    usar(escolhida)

            with ui.row().classes("gap-2"):
                ui.button("Procurar...", icon="folder_open", on_click=procurar).props(
                    "flat dense").mark(f"{chave}-procurar")  # fmt: skip
                ui.button("Usar esta pasta", icon="check",
                          on_click=lambda: usar(Path((campo.value or "").strip().strip('"'))),
                          ).props("flat dense").mark(f"{chave}-usar")  # fmt: skip
                ui.button("Voltar à padrão", icon="undo", on_click=lambda: usar(None)).props(
                    "flat dense")  # fmt: skip
