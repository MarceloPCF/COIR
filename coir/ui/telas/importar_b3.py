"""Importação das planilhas da Área do Investidor da B3 (movimentação, negociação e relatórios
consolidados), guardadas na pasta da B3 (``<dados>/_relatorios-b3`` ou a escolhida)."""

from collections.abc import Callable, Sequence
from pathlib import Path

from nicegui import events, run, ui

from coir.erros import CoirError
from coir.servicos.b3 import RelatorioB3, importar_planilhas
from coir.servicos.estado import Estado
from coir.servicos.importacao import copiar_para_pasta
from coir.ui import pastas

TIPOS = {"movimentacao": "movimentação", "negociacao": "negociação", "posicao": "posição"}


def _planilhas(pasta: Path) -> list[Path]:
    if not pasta.is_dir():
        return []
    return sorted(p for p in pasta.glob("*.xlsx") if not p.name.startswith("~$"))


def area(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    pasta_b3 = estado.pasta_b3
    ocupado = {"sim": False}

    async def importar(arquivos: Sequence[Path]) -> None:
        if ocupado["sim"] or estado.titular is None:
            if estado.titular is None:
                ui.notify("Cadastre ou escolha o titular primeiro.", type="warning")
            return
        if not arquivos:
            ui.notify("Nenhuma planilha .xlsx para importar.", type="warning")
            return
        ocupado["sim"] = True
        espera = ui.notification("Importando as planilhas da B3...", spinner=True, timeout=None)
        try:
            rel = await run.io_bound(
                importar_planilhas, estado.banco, list(arquivos), estado.titular, estado.agora()
            )
        except CoirError as erro:
            ui.notify(
                f"Importação cancelada, nada foi gravado: {erro}",
                type="negative",
                multi_line=True,
                timeout=0,
                close_button="OK",
            )
            return
        finally:
            ocupado["sim"] = False
            espera.dismiss()
        if rel is None:
            return
        estado.invalidar()
        ao_mudar()
        mostrar(rel)

    async def ao_enviar(e: events.MultiUploadEventArguments) -> None:
        recebidos = [(f.name, await f.read()) for f in e.files]
        envio.reset()
        try:
            copias = await run.io_bound(copiar_para_pasta, recebidos, pasta_b3, ".xlsx")
        except CoirError as erro:
            ui.notify(str(erro), type="negative")
            return
        await importar(copias or [])

    with ui.column().classes("coir-folha w-full gap-3"):
        ui.label(
            "Na Área do Investidor da B3, exporte em Excel a Movimentação, a Negociação e os "
            "relatórios consolidados (anual de cada ano e o mensal mais recente). Os proventos "
            "e os eventos vêm da movimentação; a posição do consolidado é a base da conciliação."
        ).classes("coir-subtitulo text-sm")
        with ui.row().classes("w-full items-center gap-3"):
            envio = (
                ui.upload(
                    multiple=True,
                    auto_upload=True,
                    on_multi_upload=ao_enviar,
                    label="Escolher planilhas (.xlsx) ou arrastar para cá",
                )
                .props('accept=".xlsx" flat')
                .classes("coir-envio min-w-[340px]")
                .mark("enviar-planilhas")
            )
            qtd = len(_planilhas(pasta_b3))
            ui.button(
                f"Importar a pasta da B3 ({qtd} planilhas)",
                icon="folder_open",
                on_click=lambda: importar(_planilhas(pasta_b3)),
            ).props("outline").mark("importar-pasta-b3")
        pastas.seletor(estado, "pasta_b3", "Pasta das planilhas", pasta_b3,
                       ui.navigate.reload)  # fmt: skip
    resultado = ui.column().classes("w-full gap-0")

    def mostrar(rel: RelatorioB3) -> None:
        resultado.clear()
        with resultado, ui.column().classes("coir-folha w-full gap-0 p-0"):
            for a in rel.arquivos:
                with ui.row().classes("coir-item-lista w-full items-start gap-3 px-4 py-2"):
                    if a.tipo is None:
                        icone = "info"  # não é planilha da B3 reconhecida, ou já importada
                    else:
                        icone = "check_circle" if a.novas else "drag_handle"
                    ui.icon(icone, size="20px").classes("coir-positivo" if a.novas else "")
                    with ui.column().classes("gap-0 grow min-w-0"):
                        ui.label(a.arquivo).classes("font-medium break-all")
                        ui.label(a.detalhe).classes("coir-subtitulo text-sm")
                    tipo = TIPOS.get(a.tipo.value, a.tipo.value) if a.tipo else "ignorado"
                    ui.label(f"{tipo} · {a.novas} nova(s) de {a.linhas}").classes(
                        "text-sm whitespace-nowrap"
                    )
