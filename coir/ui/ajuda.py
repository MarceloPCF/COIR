"""Botão "Como faço?": a página do manual da tela, ao lado dela, sem cobrir o que se está fazendo.

O texto é o de ``docs/manual/telas`` (um só para o site e para o programa); aqui saem a captura de
tela e os links, que só fazem sentido no site.
"""

import re
from pathlib import Path

from nicegui import ui

from coir.recursos import pasta_manual

RE_LINK = re.compile(r"\[([^\]]+)\]\((?!https?:)[^)]*\)")


def arquivo_da_tela(rota: str) -> Path:
    """A raiz é o Painel; as outras telas têm o nome da rota (/outros-custos, /darf...)."""
    return pasta_manual() / f"{rota.strip('/') or 'painel'}.md"


def texto_da_tela(conteudo: str) -> tuple[str, str]:
    """(título, corpo em Markdown) de uma página do manual."""
    linhas = [
        linha for linha in conteudo.splitlines()
        if not linha.startswith("![") and "Voltar ao índice" not in linha
    ]  # fmt: skip
    titulo = next((x.removeprefix("# ") for x in linhas if x.startswith("# ")), "")
    corpo = "\n".join(x for x in linhas if not x.startswith("# ")).strip()
    return titulo, RE_LINK.sub(r"\1", corpo)


def botao(rota: str) -> None:
    """O botão, no canto da tela, e a página do manual, que abre à esquerda e sem bloquear a tela:
    o formulário fica à direita, e dá para seguir os passos preenchendo."""
    folha = "coir-manual w-[420px] max-w-full h-screen no-shadow gap-2"
    with ui.dialog().props("position=left seamless") as janela, ui.card().classes(folha):
        with ui.row().classes("w-full items-center no-wrap"):
            titulo = ui.label("").classes("coir-titulo text-lg").mark("ajuda-titulo")
            ui.space()
            ui.button(icon="close", on_click=janela.close).props(
                "flat round dense").tooltip("Fechar").mark("ajuda-fechar")  # fmt: skip
        corpo = ui.markdown("").classes("w-full overflow-auto").mark("ajuda-texto")

    def abrir() -> None:
        caminho = arquivo_da_tela(rota)
        try:
            conteudo = caminho.read_text(encoding="utf-8")
        except OSError as erro:
            ui.notify(f"Manual não encontrado em {caminho}: {erro}", type="warning")
            return
        titulo_tela, texto = texto_da_tela(conteudo)
        titulo.set_text(titulo_tela)
        corpo.set_content(texto)
        janela.open()

    ui.button("Como faço?", icon="help_outline", on_click=abrir).props("flat dense").classes(
        "coir-ajuda").mark("como-faco")  # fmt: skip
