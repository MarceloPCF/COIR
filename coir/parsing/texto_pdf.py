"""Texto de cada página de um PDF (camada de texto; as notas SINACOR não precisam de OCR).

Também lê o texto já extraído de uma nota (``.txt``, páginas separadas por ``\f``, como as
fixtures e as cópias anonimizadas de ``coir notas anonimizar``): quem desenvolve roda o leitor
sobre o mesmo texto que o PDF original daria, sem ter o PDF.
"""

from collections.abc import Iterable, Sequence
from pathlib import Path
from threading import RLock

import pypdfium2 as pdfium
import pypdfium2.raw as pdfium_c

from coir.domain.cpf import so_digitos
from coir.erros import CoirError


class PdfError(CoirError):
    """PDF ilegível ou sem camada de texto."""


class PdfSenhaError(PdfError):
    """PDF protegido por senha, e nenhuma das senhas tentadas abriu."""


# Notas enviadas por e-mail costumam vir com senha formada pelo começo do CPF do titular (os 3
# primeiros dígitos é o mais comum); tentar outros começos e o CPF inteiro não custa nada.
DIGITOS_DO_CPF = (3, 4, 5, 6, 11)


SUFIXO_TEXTO = ".txt"
SEPARADOR_PAGINAS = ""

# O PDFium não é thread-safe: duas threads usando a biblioteca ao mesmo tempo corrompem o heap
# (queda nativa 0xc0000374). Todo uso do pypdfium2 (abrir, ler, desenhar, fechar) passa por aqui.
TRAVA_PDFIUM = RLock()


def senhas_do_cpf(cpfs: Iterable[str]) -> list[str]:
    """Senhas prováveis de PDF de nota a partir dos CPFs dos titulares (sem repetir)."""
    saida: list[str] = []
    for cpf in cpfs:
        digitos = so_digitos(cpf)
        if len(digitos) != 11:
            continue
        saida += [digitos[:n] for n in DIGITOS_DO_CPF if digitos[:n] not in saida]
    return saida


def abrir(caminho: Path, senhas: Sequence[str] = ()) -> pdfium.PdfDocument:
    """Abre o PDF (chamar com ``TRAVA_PDFIUM``); se for protegido, tenta cada uma das
    ``senhas`` na ordem. Nenhuma serve: ``PdfSenhaError``."""
    try:
        return pdfium.PdfDocument(caminho)
    except pdfium.PdfiumError as erro:
        if erro.err_code != pdfium_c.FPDF_ERR_PASSWORD:
            raise PdfError(f"PDF ilegível: {caminho.name}") from erro
    for senha in dict.fromkeys(s for s in senhas if s):
        try:
            return pdfium.PdfDocument(caminho, password=senha)
        except pdfium.PdfiumError as erro:
            if erro.err_code != pdfium_c.FPDF_ERR_PASSWORD:
                raise PdfError(f"PDF ilegível: {caminho.name}") from erro
    raise PdfSenhaError(
        f"PDF protegido por senha: {caminho.name}. Informe a senha do PDF ao importar (campo "
        '"Senha dos PDFs" na tela Importar); nas notas enviadas por e-mail ela costuma ser o '
        "começo do CPF"
    )


def paginas(caminho: Path, senhas: Sequence[str] = ()) -> list[str]:
    """Devolve o texto de cada página, na ordem (``senhas``: para PDF protegido)."""
    if not caminho.is_file():
        raise PdfError(f"PDF não encontrado: {caminho.name}")
    if caminho.suffix.lower() == SUFIXO_TEXTO:
        return ler_texto(caminho)
    with TRAVA_PDFIUM:
        documento = abrir(caminho, senhas)
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
