"""Texto de cada página de um PDF (camada de texto; as notas SINACOR não precisam de OCR)."""

from pathlib import Path

import pypdfium2 as pdfium

from coir.erros import CoirError


class PdfError(CoirError):
    """PDF ilegível ou sem camada de texto."""


def paginas(caminho: Path) -> list[str]:
    """Devolve o texto de cada página, na ordem."""
    if not caminho.is_file():
        raise PdfError(f"PDF não encontrado: {caminho.name}")
    try:
        documento = pdfium.PdfDocument(caminho)
    except pdfium.PdfiumError as erro:
        raise PdfError(f"PDF ilegível: {caminho.name}") from erro
    try:
        textos = [documento[i].get_textpage().get_text_range() for i in range(len(documento))]
    finally:
        documento.close()
    if not any(t.strip() for t in textos):
        raise PdfError(f"PDF sem camada de texto: {caminho.name}")
    return textos
