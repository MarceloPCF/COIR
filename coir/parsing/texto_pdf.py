"""Texto de cada página de um PDF (camada de texto; as notas SINACOR não precisam de OCR).

Também lê o texto já extraído de uma nota (``.txt``, páginas separadas por ``\f``, como as
fixtures e as cópias anonimizadas de ``coir notas anonimizar``): quem desenvolve roda o leitor
sobre o mesmo texto que o PDF original daria, sem ter o PDF.
"""

from pathlib import Path
from threading import RLock

import pypdfium2 as pdfium

from coir.erros import CoirError


class PdfError(CoirError):
    """PDF ilegível ou sem camada de texto."""


SUFIXO_TEXTO = ".txt"
SEPARADOR_PAGINAS = ""

# O PDFium não é thread-safe: duas threads usando a biblioteca ao mesmo tempo corrompem o heap
# (queda nativa 0xc0000374). Todo uso do pypdfium2 (abrir, ler, desenhar, fechar) passa por aqui.
TRAVA_PDFIUM = RLock()


def paginas(caminho: Path) -> list[str]:
    """Devolve o texto de cada página, na ordem."""
    if not caminho.is_file():
        raise PdfError(f"PDF não encontrado: {caminho.name}")
    if caminho.suffix.lower() == SUFIXO_TEXTO:
        return ler_texto(caminho)
    with TRAVA_PDFIUM:
        try:
            documento = pdfium.PdfDocument(caminho)
        except pdfium.PdfiumError as erro:
            raise PdfError(f"PDF ilegível: {caminho.name}") from erro
        try:
            textos = [_texto(documento[i]) for i in range(len(documento))]
        finally:
            documento.close()
    if not any(t.strip() for t in textos):
        raise PdfError(
            f"PDF sem texto, só imagem (foto ou digitalização): {caminho.name}. Baixe a nota em "
            "PDF no site da corretora; o COIR não lê notas fotografadas ou escaneadas"
        )
    return textos


def _texto(pagina: pdfium.PdfPage) -> str:
    """Texto de uma página, liberando a página e o textpage ainda dentro da trava (o coletor de
    lixo poderia liberá-los depois, em outra thread, sem ela)."""
    try:
        textpage = pagina.get_textpage()
        try:
            return textpage.get_text_range()
        finally:
            textpage.close()
    finally:
        pagina.close()


def ler_texto(caminho: Path) -> list[str]:
    """Páginas de um texto de nota gravado com ``gravar_texto`` (fins de linha preservados)."""
    try:
        with caminho.open(encoding="utf-8", newline="") as arquivo:
            textos = arquivo.read().split(SEPARADOR_PAGINAS)
    except (OSError, UnicodeDecodeError) as erro:
        raise PdfError(f"texto de nota ilegível: {caminho.name}") from erro
    if not any(t.strip() for t in textos):
        raise PdfError(f"texto de nota vazio: {caminho.name}")
    return textos


def gravar_texto(caminho: Path, textos: list[str]) -> None:
    """Grava as páginas separadas por ``\f``, sem mexer nos fins de linha (``\r\n`` do
    pdfium): lido de volta, dá exatamente o mesmo texto."""
    if any(SEPARADOR_PAGINAS in t for t in textos):
        raise PdfError("página com o separador de páginas no texto")
    with caminho.open("w", encoding="utf-8", newline="") as arquivo:
        arquivo.write(SEPARADOR_PAGINAS.join(textos))
