"""Tela ainda não construída: título, o que ela vai mostrar e o aviso."""

from nicegui import ui


def montar(titulo: str, descricao: str) -> None:
    ui.label(titulo).classes("coir-titulo-pagina")
    ui.label(descricao).classes("coir-subtitulo")
    ui.label("Esta tela ainda está em construção.").classes("coir-vazio w-full")
