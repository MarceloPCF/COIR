"""Tela 3 — Importar notas: escolher PDFs (ou a pasta das notas), progresso e o resultado
arquivo a arquivo. Os PDFs escolhidos são guardados na pasta das notas (para a conferência).
"""

from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from pathlib import Path

from nicegui import events, run, ui

from coir.erros import CoirError
from coir.servicos.estado import Estado
from coir.servicos.importacao import (
    RelatorioImportacao,
    ResultadoArquivo,
    copiar_para_pasta,
    divergencias_conhecidas,
    importar_lista,
    por_arquivo,
)
from coir.ui.telas import importar_b3

ICONES = {
    "gravado": ("check_circle", "coir-positivo"),
    "repetido": ("drag_handle", ""),
    "rejeitado": ("cancel", "coir-negativo"),
    "aviso": ("warning", "coir-carimbo"),
}
ROTULOS = {
    "gravado": "gravado",
    "repetido": "já importado",
    "rejeitado": "rejeitado",
    "aviso": "gravado com aviso",
}


@dataclass
class _Andamento:
    feitos: int = 0
    total: int = 0
    arquivo: str = ""
    ativo: bool = False
    erros: list[str] = field(default_factory=list[str])

    def __call__(self, feitos: int, total: int, arquivo: str) -> None:  # chamado na thread
        self.feitos, self.total, self.arquivo = feitos, total, arquivo


def _pdfs(pasta: Path) -> list[Path]:
    if not pasta.is_dir():
        return []
    return sorted(p for p in pasta.iterdir() if p.is_file() and p.suffix.lower() == ".pdf")


def area(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    """Escolha dos arquivos, andamento e resultado (também usada no assistente)."""
    andamento = _Andamento()

    async def importar(arquivos: Sequence[Path]) -> None:
        if andamento.ativo:
            return
        if not arquivos:
            ui.notify("Nenhum PDF para importar.", type="warning")
            return
        andamento.ativo, andamento.erros = True, []
        andamento(0, len(arquivos), arquivos[0].name)
        travar(True)
        try:
            rel = await run.io_bound(
                importar_lista,
                estado.banco,
                arquivos,
                estado.refs,
                estado.agora(),
                None,
                estado.extrator,
                andamento,
            )
        except CoirError as erro:  # lote desfeito: nada gravado
            ui.notify(
                f"Importação cancelada, nada foi gravado: {erro}",
                type="negative",
                multi_line=True,
                timeout=0,
                close_button="OK",
            )
            return
        finally:
            andamento.ativo = False
            travar(False)
            mostrar_andamento()
        if rel is None:
            return
        estado.invalidar()
        mostrar(rel)
        if rel.mudou or rel.titulares_criados:
            ao_mudar()

    async def ao_enviar(e: events.MultiUploadEventArguments) -> None:
        recebidos = [(f.name, await f.read()) for f in e.files]
        envio.reset()
        try:
            copias = await run.io_bound(copiar_para_pasta, recebidos, estado.pasta_notas)
        except CoirError as erro:
            ui.notify(str(erro), type="negative")
            return
        await importar(copias or [])

    async def importar_pasta() -> None:
        await importar(_pdfs(estado.pasta_notas))

    with ui.column().classes("coir-folha w-full gap-3"):
        with ui.row().classes("w-full items-center gap-3"):
            envio = (
                ui.upload(
                    multiple=True,
                    auto_upload=True,
                    on_multi_upload=ao_enviar,
                    label="Escolher PDFs ou arrastar para cá",
                )
                .props('accept=".pdf,application/pdf" flat')
                .classes("coir-envio min-w-[340px]")
                .mark("enviar-pdfs")
            )
            qtd = len(_pdfs(estado.pasta_notas))
            pasta = (
                ui.button(
                    f"Importar a pasta das notas ({qtd} PDFs)",
                    icon="folder_open",
                    on_click=importar_pasta,
                )
                .props("outline")
                .mark("importar-pasta")
            )
        ui.label(f"Pasta das notas: {estado.pasta_notas}").classes("coir-subtitulo text-xs")
        barra = ui.linear_progress(value=0, show_value=False).props("rounded size=8px")
        texto = ui.label("").classes("coir-subtitulo text-sm")

    def travar(sim: bool) -> None:
        envio.set_enabled(not sim)
        pasta.set_enabled(not sim)

    def mostrar_andamento() -> None:
        visivel = andamento.ativo and andamento.total > 0
        barra.set_visibility(visivel)
        texto.set_visibility(visivel)
        if visivel:
            barra.set_value(andamento.feitos / andamento.total)
            lendo = f": {andamento.arquivo}" if andamento.arquivo else ": gravando"
            texto.set_text(f"{andamento.feitos} de {andamento.total} arquivos{lendo}")

    mostrar_andamento()
    ui.timer(0.25, mostrar_andamento)
    resultado = ui.column().classes("w-full gap-2")

    def mostrar(rel: RelatorioImportacao) -> None:
        linhas = por_arquivo(rel, divergencias_conhecidas(estado.dados))
        resultado.clear()
        with resultado:
            _resumo(rel, linhas)
            with ui.column().classes("coir-folha w-full gap-0 p-0"):
                for linha in linhas:
                    _linha(linha)
            for aviso in (a for a in rel.avisos if not any(x.arquivo in a for x in linhas)):
                ui.label(f"Aviso: {aviso}").classes("coir-carimbo text-sm")


def _resumo(rel: RelatorioImportacao, linhas: list[ResultadoArquivo]) -> None:
    rejeitados = sum(1 for x in linhas if x.situacao == "rejeitado")
    with ui.row().classes("items-baseline gap-6"):
        for rotulo, valor, classe in (
            ("arquivos", len(linhas), ""),
            ("notas gravadas", rel.notas_gravadas, "coir-positivo"),
            ("notas que já estavam no banco", rel.notas_ja_no_banco, ""),
            ("arquivos já importados", len(rel.arquivos_ignorados), ""),
            ("notas rejeitadas", len(rel.rejeitadas), "coir-negativo" if rejeitados else ""),
        ):
            with ui.column().classes("gap-0"):
                ui.label(str(valor)).classes(f"coir-numero text-2xl {classe}")
                ui.label(rotulo).classes("coir-campo-rotulo")
    if rel.titulares_criados:
        ui.label("Titular criado pela nota: confira o nome em Cadastros.").classes("coir-carimbo")
    if rel.notas_renormalizadas:
        ui.label(
            f"{rel.notas_renormalizadas} nota(s) já gravadas tiveram as operações regravadas "
            "(tabelas de referência mudaram)."
        ).classes("coir-subtitulo")


def _linha(x: ResultadoArquivo) -> None:
    icone, classe = ICONES[x.situacao]
    with ui.row().classes("coir-item-lista w-full items-start gap-3 px-4 py-2"):
        ui.icon(icone, size="20px").classes(classe)
        with ui.column().classes("gap-0 grow min-w-0"):
            ui.label(x.arquivo).classes("font-medium break-all")
            for d in x.detalhes:
                ui.label(d).classes("coir-subtitulo text-sm")
        situacao = ROTULOS[x.situacao] + (" (divergência conhecida)" if x.conhecida else "")
        if x.notas:
            situacao = f"{x.notas} nota{'s' if x.notas > 1 else ''} · {situacao}"
        ui.label(situacao).classes(f"text-sm whitespace-nowrap {classe}")


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Importar").classes("coir-titulo-pagina")
    with ui.tabs().props("no-caps align=left").classes("w-full") as abas:
        t_notas = ui.tab("Notas de corretagem (PDF)").mark("aba-notas")
        t_b3 = ui.tab("Planilhas da B3 (xlsx)").mark("aba-b3")
    with ui.tab_panels(abas, value=t_notas).classes("w-full bg-transparent"):
        with ui.tab_panel(t_notas).classes("p-0 pt-3 gap-3"):
            ui.label(
                "Escolha os PDFs das notas de corretagem. O COIR confere cada nota (negócios, "
                "custos e líquido) antes de gravar; a que não fecha ao centavo não entra, e o "
                "motivo aparece abaixo. Importar de novo o mesmo arquivo não duplica nada.",
            ).classes("coir-subtitulo")
            area(estado, ao_mudar)
        with ui.tab_panel(t_b3).classes("p-0 pt-3"):
            importar_b3.area(estado, ao_mudar)
