"""Tela Suporte: os dados que cada titular quer ocultar e o envio de notas para quem desenvolve
(texto anonimizado e imagens com tarja, conferidos pelo usuário antes de sair, por e-mail)."""

import os
import sys
import webbrowser
from collections.abc import Callable
from pathlib import Path

from nicegui import events, run, ui

from coir.config import VERSAO
from coir.domain.cpf import cpf_valido, so_digitos
from coir.erros import CoirError
from coir.servicos import suporte
from coir.servicos.estado import Estado
from coir.servicos.suporte import CAMPOS_LIVRES, DadosPessoais

NOVO = "novo"
ROTULOS = {
    "endereco": "Endereço (como sai nas notas; um por linha se mudou)",
    "bairro": "Bairro",
    "cidade": "Cidade",
    "cep": "CEP",
    "telefones": "Telefones (um por linha)",
    "email": "E-mail",
    "assessor": "Assessor (nome ou código)",
    "banco": "Banco, agência e conta corrente",
    "outros": "Outros termos a ocultar (um por linha)",
}


def _secao(titulo: str, texto: str) -> ui.column:
    with ui.column().classes("coir-folha w-full gap-2") as coluna:
        ui.label(titulo).classes("coir-titulo text-base")
        ui.label(texto).classes("coir-subtitulo text-sm")
    return coluna


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Suporte").classes("coir-titulo-pagina")
    ui.label(
        "Para uma nota que o COIR não lê ou lê errado (uma corretora nova, um layout diferente)."
    ).classes("coir-subtitulo")
    _dados(estado)
    _envio(estado)


# --------------------------------------------------------------------------- dados para ocultar
def _dados(estado: Estado) -> None:
    with _secao(
        "Dados para ocultar",
        "O COIR troca por fictícios, nas notas enviadas, os dados de cada titular informados "
        "aqui, onde quer que apareçam, em qualquer corretora. Nome, CPF e contas vêm do cadastro; "
        "complete com o resto, como sai nas notas. Ficam só neste computador (pasta de dados, "
        "_privado) e não vão no envio.",
    ):
        escolha = ui.select({}, label="Titular").props("outlined dense").classes("w-[360px]")
        escolha.mark("ocultar-titular")
        with ui.grid(columns=2).classes("w-full max-w-[900px] gap-2"):
            nome = (
                ui.input("Nome (como sai nas notas)").props("outlined dense").mark("ocultar-nome")
            )
            cpf = ui.input("CPF", validation={"CPF inválido": lambda v: cpf_valido(v or "")})
            cpf.props("outlined dense").mark("ocultar-cpf")
            contas = ui.input("Contas nas corretoras (separadas por vírgula)").props(
                "outlined dense").mark("ocultar-contas")  # fmt: skip
            campos = {c: ui.textarea(ROTULOS[c]).props("outlined dense autogrow").mark(
                f"ocultar-{c}") for c in CAMPOS_LIVRES}  # fmt: skip
        ui.button("Gravar", icon="save", on_click=lambda: gravar()).mark("ocultar-gravar")

    lista: dict[str, DadosPessoais] = {}

    def ler(manter: str | None = None) -> None:
        try:
            todos = suporte.dados_para_ocultar(estado.dados, estado.banco)
        except CoirError as erro:
            ui.notify(str(erro), type="negative")
            todos = []
        lista.clear()
        lista.update({so_digitos(d.cpf): d for d in todos})
        opcoes = {c: d.nome or "(sem nome)" for c, d in lista.items()} | {NOVO: "Outro titular"}
        escolha.set_options(opcoes, value=manter if manter in opcoes else next(iter(opcoes)))

    def mostrar() -> None:
        d = lista.get(escolha.value or "") or DadosPessoais(cpf="")
        nome.value, cpf.value, contas.value = d.nome, d.cpf, ", ".join(d.contas)
        for c, campo in campos.items():
            campo.value = getattr(d, c)

    def gravar() -> None:
        try:
            salvo = suporte.salvar_titular(estado.dados, DadosPessoais(
                cpf=cpf.value or "", nome=nome.value or "",
                contas=[c for c in (contas.value or "").split(",")],
                **{c: campo.value or "" for c, campo in campos.items()},
            ))  # fmt: skip
        except CoirError as erro:
            ui.notify(str(erro), type="negative")
            return
        ui.notify("Dados para ocultar gravados.", type="positive")
        ler(salvo.cpf)

    escolha.on_value_change(lambda _: mostrar())
    ler()
    mostrar()


# ------------------------------------------------------------------------------------- envio
def _envio(estado: Estado) -> None:
    with _secao(
        "Enviar notas para suporte",
        "Escolha os PDFs das notas. O COIR grava numa pasta o texto delas e a imagem de cada "
        "página, com nome, CPF, conta, endereço, telefone, e-mail, assessor e dados bancários "
        "trocados por fictícios (nas imagens, cobertos por tarja preta). A cópia só sai se o "
        "COIR ler nela o mesmo que lê no original e não sobrar nenhum dado pessoal. Atenção: as "
        "notas continuam mostrando as operações (ativos, datas, quantidades e valores). Confira "
        "as imagens antes de enviar; os PDFs não são guardados.",
    ):
        with ui.row().classes("items-center gap-3"):
            envio = (
                ui.upload(
                    multiple=True,
                    auto_upload=True,
                    on_multi_upload=lambda e: preparar(e),
                    label="Escolher os PDFs das notas",
                )
                .props('accept=".pdf,application/pdf" flat')
                .classes("coir-envio min-w-[300px]")
                .mark("suporte-pdfs")
            )
            ui.button("Abrir a pasta", icon="folder_open",
                      on_click=lambda: _abrir_pasta(estado.dados / suporte.PASTA_SUPORTE),
                      ).props("flat dense")  # fmt: skip
        resultado = ui.column().classes("w-full gap-2").mark("suporte-resultado")

    async def preparar(e: events.MultiUploadEventArguments) -> None:
        recebidos = [(f.name, await f.read()) for f in e.files]
        envio.reset()
        resultado.clear()
        try:
            pronto = await run.io_bound(
                suporte.preparar_envio, estado.dados, estado.banco, recebidos, estado.agora()
            )
        except CoirError as erro:
            ui.notify(str(erro), type="negative", multi_line=True)
            return
        if pronto is None:  # run.io_bound devolve None se a página fechou
            return
        with resultado:
            _conferencia(pronto)


def _conferencia(envio: suporte.Envio) -> None:
    """Uma linha por nota (titular, páginas, trechos ocultados, situação) e, se alguma passou, a
    confirmação de que as imagens foram conferidas, o .zip e o e-mail."""
    with ui.row().classes("items-center gap-3"):
        ui.label(f"Preparado em {envio.pasta}").classes("coir-numero text-sm")
        ui.button("Abrir a pasta", icon="folder_open",
                  on_click=lambda: _abrir_pasta(envio.pasta)).props("flat dense")  # fmt: skip
    for nota in envio.notas:
        with ui.row().classes("coir-item-lista w-full items-center gap-4 py-1"):
            ui.label(nota.origem).classes("text-sm w-[260px] truncate")
            ui.label(envio.titular(nota)).classes("coir-subtitulo text-sm w-[220px]")
            if nota.destino is None:
                ui.label("não anonimizada, não vai no envio").classes("coir-negativo text-sm")
                continue
            ui.label(f"{nota.destino.stem}: {nota.paginas} página(s), "
                     f"{nota.ocultados} trecho(s) ocultado(s)").classes("text-sm")  # fmt: skip
            if nota.alertas:
                ui.label(f"{nota.alertas} trecho(s) dos seus dados à mostra (contorno laranja): "
                         "confira").classes("coir-negativo text-sm")  # fmt: skip
            ui.space()
            if nota.imagens:
                ui.button("Ver", icon="visibility", on_click=_ver(nota)).props("flat dense").mark(
                    f"ver-{nota.destino.stem}")  # fmt: skip
        for motivo in nota.recusas:
            ui.label(f"· {motivo}").classes("coir-subtitulo text-xs")
    if not envio.aceitas:
        return
    with ui.row().classes("items-center gap-3"):
        conferi = ui.checkbox("Conferi as imagens: nenhum dado pessoal ficou à mostra").mark(
            "suporte-conferi"
        )
        gerar = ui.button("Gravar o .zip", icon="archive", on_click=lambda: gravar()).mark(
            "suporte-zip")  # fmt: skip
        gerar.bind_enabled_from(conferi, "value")
    final = ui.column().classes("w-full gap-1").mark("suporte-final")

    def gravar() -> None:
        try:
            zip_ = suporte.gerar_zip(envio)
        except CoirError as erro:
            ui.notify(str(erro), type="negative")
            return
        final.clear()
        with final:
            ui.label(f"Gravado: {zip_}").classes("coir-positivo text-sm")
            ui.label(
                f"Envie para {suporte.EMAIL_SUPORTE}: o botão abre a mensagem no seu programa de "
                "e-mail; anexe o .zip, descreva o problema e envie."
            ).classes("text-sm")
            with ui.row().classes("gap-3"):
                ui.button("Escrever e-mail", icon="mail",
                          on_click=lambda: webbrowser.open(suporte.endereco_email(zip_, VERSAO)),
                          ).mark("suporte-email")  # fmt: skip
                ui.button("Abrir a pasta", icon="folder_open",
                          on_click=lambda: _abrir_pasta(envio.pasta)).props("flat")  # fmt: skip


def _ver(nota: suporte.ResultadoArquivo) -> Callable[[], None]:
    def abrir() -> None:
        with ui.dialog() as janela, ui.card().classes("w-[1000px] max-w-full"):
            ui.label(f"{nota.origem} — confira as tarjas").classes("coir-titulo text-base")
            with ui.scroll_area().classes("w-full h-[70vh]"):
                for imagem in nota.imagens:
                    ui.image(imagem).classes("w-full border")
            ui.button("Fechar", on_click=janela.close).props("flat")
        janela.open()

    return abrir


def _abrir_pasta(pasta: Path) -> None:
    if sys.platform == "win32" and pasta.is_dir():
        os.startfile(pasta)
    else:
        ui.notify(str(pasta))
