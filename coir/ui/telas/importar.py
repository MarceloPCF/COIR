"""Tela 3 — Importar notas: escolher PDFs (ou a pasta das notas), progresso e o resultado
arquivo a arquivo. Os PDFs escolhidos são guardados na pasta das notas (para a conferência).
"""

from collections.abc import Awaitable, Callable, Sequence
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from nicegui import events, run, ui

from coir.erros import CoirError
from coir.servicos.estado import Estado
from coir.servicos.importacao import (
    NomeAmbiguo,
    RelatorioImportacao,
    ResultadoArquivo,
    copiar_para_pasta,
    divergencias_conhecidas,
    escolher_codigo,
    importar_lista,
    por_arquivo,
)
from coir.ui import ocupado, pastas
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
        ocupado.detalhar(f"{feitos} de {total} arquivos" + ("" if arquivo else ": gravando"))


def _pdfs(pasta: Path) -> list[Path]:
    if not pasta.is_dir():
        return []
    return sorted(p for p in pasta.iterdir() if p.is_file() and p.suffix.lower() == ".pdf")


@dataclass
class _Controles:
    envio: ui.upload
    pasta: ui.button
    senha: ui.input
    reprocessar: ui.checkbox
    barra: ui.linear_progress
    texto: ui.label


def _controles(
    estado: Estado,
    ao_enviar: Callable[[events.MultiUploadEventArguments], Awaitable[None]],
    importar_pasta: Callable[[], Awaitable[None]],
) -> _Controles:
    """Envio dos PDFs, pasta das notas, senha, reprocessar e a barra de andamento."""
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
        senha = (
            ui.input(
                "Senha dos PDFs (se houver)",
                password=True,
                password_toggle_button=True,
            )
            .props("outlined dense")
            .classes("min-w-[340px]")
            .tooltip(
                "O COIR já tenta o começo do CPF dos titulares cadastrados. A senha vale para "
                "todos os PDFs do lote e fica só na memória até fechar o programa."
            )
            .mark("senha-pdfs")
        )
        reprocessar = (
            ui.checkbox("Reprocessar notas já importadas")
            .tooltip(
                "Lê de novo os PDFs escolhidos e fica com a releitura, útil se o COIR leu algo "
                "errado ou se a corretora corrigiu a nota. Faz backup antes; nota que a "
                "releitura não confirma continua como está."
            )
            .mark("reprocessar")
        )
        pastas.seletor(estado, "pasta_notas", "Pasta das notas", estado.pasta_notas,
                       ui.navigate.reload)  # fmt: skip
        barra = ui.linear_progress(value=0, show_value=False).props("rounded size=8px")
        texto = ui.label("").classes("coir-subtitulo text-sm")
    return _Controles(envio, pasta, senha, reprocessar, barra, texto)


async def _eventos_automaticos(estado: Estado) -> None:
    """Depois de importar: lança os eventos seguros que as notas novas tornaram aplicáveis."""
    try:
        with ocupado.processando("Conferindo os eventos das notas importadas"):
            automaticos = await run.io_bound(estado.lancar_eventos_automaticos, True)
    except CoirError as erro:
        ui.notify(f"Eventos não lançados: {erro}", type="warning", multi_line=True)
        return
    if automaticos and automaticos.ids:
        ui.notify(f"{len(automaticos.ids)} evento(s) lançado(s) automaticamente. Veja na tela "
                  "Eventos.", type="positive", multi_line=True)  # fmt: skip


def area(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    """Escolha dos arquivos, andamento e resultado (também usada no assistente)."""
    andamento = _Andamento()
    lidos: list[Path] = []  # os arquivos da última importação (para importar de novo os recusados)

    async def gravar_escolhas(escolhas: dict[NomeAmbiguo, str], arquivos: set[str]) -> None:
        try:
            for a, ticker in escolhas.items():
                await run.io_bound(escolher_codigo, estado.banco, a.corretora, a.conta, a.nome,
                                   a.data, ticker, estado.agora())  # fmt: skip
        except CoirError as erro:
            ui.notify(f"Escolha não gravada: {erro}", type="negative", multi_line=True)
            return
        await importar([p for p in lidos if p.name in arquivos])

    async def importar(arquivos: Sequence[Path]) -> None:
        if andamento.ativo:
            return
        if not arquivos:
            ui.notify("Nenhum PDF para importar.", type="warning")
            return
        digitada = str(c.senha.value or "").strip()
        if digitada and digitada not in estado.senhas_informadas:
            estado.senhas_informadas.append(digitada)  # vale até fechar o COIR; nunca gravada
        andamento.ativo, andamento.erros = True, []
        lidos[:] = arquivos
        andamento(0, len(arquivos), arquivos[0].name)
        travar(True)
        try:
            with ocupado.processando("Importando as notas"):
                rel = await run.io_bound(
                    importar_lista,
                    estado.banco,
                    arquivos,
                    estado.refs,
                    estado.agora(),
                    None,
                    estado.extrator,
                    andamento,
                    list(estado.senhas_informadas),
                    bool(c.reprocessar.value),
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
        if rel.mudou:  # as notas novas podem pôr ativos em carteira na data de um evento
            await _eventos_automaticos(estado)
        if rel.mudou or rel.titulares_criados:
            ao_mudar()

    async def ao_enviar(e: events.MultiUploadEventArguments) -> None:
        recebidos = [(f.name, await f.read()) for f in e.files]
        if not c.envio.is_deleted:  # a tela pode ter sido refeita durante o envio
            c.envio.reset()
        try:
            copias = await run.io_bound(copiar_para_pasta, recebidos, estado.pasta_notas)
        except CoirError as erro:
            ui.notify(str(erro), type="negative")
            return
        await importar(copias or [])

    async def importar_pasta() -> None:
        await importar(_pdfs(estado.pasta_notas))

    c = _controles(estado, ao_enviar, importar_pasta)

    def travar(sim: bool) -> None:
        if c.envio.is_deleted:
            return  # a importação continua, mas o usuário saiu da tela (ou ela foi refeita)
        c.envio.set_enabled(not sim)
        c.pasta.set_enabled(not sim)

    def mostrar_andamento() -> None:
        if c.barra.is_deleted:
            return
        visivel = andamento.ativo and andamento.total > 0
        c.barra.set_visibility(visivel)
        c.texto.set_visibility(visivel)
        if visivel:
            c.barra.set_value(andamento.feitos / andamento.total)
            lendo = f": {andamento.arquivo}" if andamento.arquivo else ": gravando"
            c.texto.set_text(f"{andamento.feitos} de {andamento.total} arquivos{lendo}")

    mostrar_andamento()
    ui.timer(0.25, mostrar_andamento)
    resultado = ui.column().classes("w-full gap-2")

    def mostrar(rel: RelatorioImportacao) -> None:
        if resultado.is_deleted:
            return  # o usuário saiu da tela (ou ela foi refeita) durante a importação
        linhas = por_arquivo(rel, divergencias_conhecidas(estado.dados))
        resultado.clear()
        with resultado:
            _resultado(rel, linhas, gravar_escolhas)


def _resultado(
    rel: RelatorioImportacao, linhas: list[ResultadoArquivo],
    gravar_escolhas: Callable[[dict[NomeAmbiguo, str], set[str]], Awaitable[None]],
) -> None:  # fmt: skip
    """O resultado da importação: resumo, códigos a escolher, cada arquivo e os avisos."""
    _resumo(rel, linhas)
    if rel.ambiguos:
        _perguntar_codigos(rel.ambiguos, gravar_escolhas)
    with ui.column().classes("coir-folha w-full gap-0 p-0"):
        for linha in linhas:
            _linha(linha)
    avisos = [a for a in rel.avisos if not any(x.arquivo in a for x in linhas)]
    for aviso in (a for a in avisos if "já importada" not in a):
        ui.label(f"Aviso: {aviso}").classes("coir-carimbo text-sm")
    repetidas = [a for a in avisos if "já importada" in a]
    if rel.arquivos_ignorados or rel.notas_ja_no_banco or repetidas:
        ui.label(
            "Notas que já estavam no banco não são gravadas de novo (cada nota entra uma "
            'vez só). Para ler de novo, marque "Reprocessar notas já importadas". Os '
            "PDFs continuam na pasta onde estão; a tela Notas abre o original de lá."
        ).classes("coir-subtitulo text-sm").mark("info-repetidas")
        for aviso in repetidas:
            ui.label(aviso).classes("coir-subtitulo text-xs")


def _resumo(rel: RelatorioImportacao, linhas: list[ResultadoArquivo]) -> None:
    rejeitados = sum(1 for x in linhas if x.situacao == "rejeitado")
    with ui.row().classes("items-baseline gap-6"):
        for rotulo, valor, classe in (
            ("arquivos", len(linhas), ""),
            ("notas gravadas", rel.notas_gravadas, "coir-positivo"),
            *(
                [("notas substituídas", rel.notas_substituidas, "coir-carimbo")]
                if rel.notas_substituidas
                else []
            ),
            *(
                [("registros antigos removidos", rel.notas_removidas, "coir-carimbo")]
                if rel.notas_removidas
                else []
            ),
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


def _perguntar_codigos(
    ambiguos: Sequence[NomeAmbiguo],
    gravar: Callable[[dict[NomeAmbiguo, str], set[str]], Awaitable[None]],
) -> None:
    """Nome de pregão que valia para mais de um código na data: pergunta qual é, uma vez por
    conta, nome e pregão (os dois códigos podem ter sido negociados em dias diferentes), e
    importa de novo as notas recusadas por isso."""
    perguntas: dict[tuple[str, str, str, date], NomeAmbiguo] = {}
    for a in ambiguos:
        perguntas.setdefault((a.corretora, a.conta, a.nome, a.data), a)
    escolhas: dict[NomeAmbiguo, ui.radio] = {}
    with ui.column().classes("coir-folha w-full gap-2").mark("codigos-em-duvida"):
        ui.label("Qual é o código do ativo?").classes("coir-titulo text-base")
        ui.label(
            "A nota traz só o nome de pregão, e na data ele valia para mais de um código. O COIR "
            "não escolhe no palpite. Confira no extrato da corretora ou na Área do Investidor "
            "da B3 e marque o código negociado em cada pregão; a escolha fica guardada para "
            "as notas dessa conta naquele dia. Se preferir, importe antes a planilha de "
            "negociação da B3 do período: com ela o COIR acha o código sozinho."
        ).classes("coir-subtitulo text-sm")
        for a in perguntas.values():
            with ui.row().classes("items-center gap-4"):
                ui.label(
                    f"{a.especificacao}, nota de {a.data:%d/%m/%Y} ({a.corretora}, conta {a.conta})"
                ).classes("font-medium")
                escolhas[a] = (
                    ui.radio(list(a.codigos))
                    .props("inline dense")
                    .mark(f"codigo-{a.nome}-{a.data:%Y%m%d}")
                )

        async def confirmar() -> None:
            marcadas = {a: str(r.value) for a, r in escolhas.items() if r.value}
            if not marcadas:
                ui.notify("Marque o código de pelo menos um ativo.", type="warning")
                return
            chaves = {(a.corretora, a.conta, a.nome, a.data) for a in marcadas}
            arquivos = {a.arquivo for a in ambiguos
                        if (a.corretora, a.conta, a.nome, a.data) in chaves}  # fmt: skip
            await gravar(marcadas, arquivos)

        ui.button("Gravar a escolha e importar de novo", icon="check", on_click=confirmar).mark(
            "gravar-codigos"
        )


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
            plural = "s" if x.notas > 1 else ""
            situacao = (f"{x.notas} nota{plural} gravada{plural}; outras recusadas"
                        if x.situacao == "rejeitado"
                        else f"{x.notas} nota{plural} · {situacao}")  # fmt: skip
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
