"""Imagem de cada página de uma nota com tarja sobre o que o anonimizador trocou.

Acompanha o texto anonimizado no pacote do suporte: quem desenvolve vê o layout (qual coluna é
preço, qual é quantidade) e testa o leitor no texto. A posição de cada caractere vem do próprio
PDF (pypdfium2), na mesma ordem do texto que o COIR lê, então a tarja cai exatamente sobre o
trecho trocado. Trecho informado pelo titular que ficou à mostra (o bairro "Vale" no ativo VALE ON)
ganha um contorno, para ele conferir.
"""

from collections.abc import Callable, Sequence
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageDraw

from coir.erros import AnonimizacaoError

DPI = 110  # legível na tela e pequeno o bastante para o e-mail (100 a 200 KB por página)
FOLGA = 1  # pixels em volta de cada tarja (borda suavizada das letras)
CORES = 16  # paleta da imagem gravada (texto preto e branco, tarja e contorno)
CONTORNO = (230, 120, 0)

Retangulo = tuple[float, float, float, float]  # esquerda, cima, direita, baixo (pixels)
Posicao = Callable[[float, float], tuple[float, float]]  # ponto do PDF -> pixel


def paginas_tarjadas(
    pdf: Path,
    textos: Sequence[str],
    trocados: Sequence[bytes],
    alertas: Sequence[Sequence[tuple[int, int]]] = (),
    dpi: int = DPI,
) -> list[Image.Image]:
    """Uma imagem por página. ``textos`` é o texto original de cada página (o que foi
    anonimizado): se o PDF der outro texto, as posições não casam e nada é gerado."""
    escala = dpi / 72
    try:
        documento = pdfium.PdfDocument(pdf)
    except pdfium.PdfiumError as erro:
        raise AnonimizacaoError(f"PDF ilegível: {pdf.name}") from erro
    try:
        if len(documento) != len(textos):
            raise AnonimizacaoError("o PDF tem outro número de páginas que o texto conferido")
        imagens = []
        for i in range(len(documento)):
            pagina = documento[i]
            textpage = pagina.get_textpage()
            if textpage.count_chars() != len(textos[i]) or textpage.get_text_range() != textos[i]:
                raise AnonimizacaoError(f"página {i + 1}: o texto do PDF não casa com o conferido")
            bitmap = pagina.render(scale=escala, grayscale=True)  # pyright: ignore[reportArgumentType]
            imagem = bitmap.to_pil().convert("RGB")
            desenho = ImageDraw.Draw(imagem)
            posicao = _posicao(pagina, escala)
            for ini, fim in _trechos(trocados[i]):
                for r in _caixas(textpage, ini, fim, posicao):
                    desenho.rectangle(_folga(r, FOLGA), fill=(0, 0, 0))
            for ini, fim in alertas[i] if i < len(alertas) else ():
                for r in _caixas(textpage, ini, fim, posicao):
                    desenho.rectangle(_folga(r, FOLGA + 2), outline=CONTORNO, width=2)
            imagens.append(imagem.quantize(colors=CORES))
        return imagens
    finally:
        documento.close()


def gravar_png(imagem: Image.Image, destino: Path) -> None:
    imagem.save(destino, format="PNG", optimize=True)


def _trechos(mascara: bytes) -> list[tuple[int, int]]:
    """Sequências de caracteres trocados (início, fim)."""
    trechos, inicio = [], None
    for i, marcado in enumerate([*mascara, 0]):
        if marcado and inicio is None:
            inicio = i
        elif not marcado and inicio is not None:
            trechos.append((inicio, i))
            inicio = None
    return trechos


def _posicao(pagina: pdfium.PdfPage, escala: float) -> Posicao:
    """Converte um ponto do PDF (x, y) na posição do pixel da imagem, com a rotação da página."""
    caixa = pagina.get_cropbox()
    if caixa is None:
        raise AnonimizacaoError("página sem área de impressão")
    esquerda, baixo, direita, cima = caixa
    largura, altura = (direita - esquerda) * escala, (cima - baixo) * escala
    rotacao = pagina.get_rotation() % 360

    def converter(x: float, y: float) -> tuple[float, float]:
        u, v = (x - esquerda) * escala, (cima - y) * escala
        if rotacao == 90:
            return altura - v, u
        if rotacao == 180:
            return largura - u, altura - v
        if rotacao == 270:
            return v, largura - u
        return u, v

    return converter


def _caixas(
    textpage: pdfium.PdfTextPage, inicio: int, fim: int, posicao: Posicao
) -> list[Retangulo]:
    """Retângulo de cada caractere do trecho (os gerados, como a quebra de linha, não têm)."""
    caixas = []
    for i in range(inicio, fim):
        esquerda, baixo, direita, cima = textpage.get_charbox(i, loose=True)
        if direita <= esquerda or cima <= baixo:
            continue
        x1, y1 = posicao(esquerda, cima)
        x2, y2 = posicao(direita, baixo)
        caixas.append((min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2)))
    return caixas


def _folga(r: Retangulo, folga: int) -> Retangulo:
    return (r[0] - folga, r[1] - folga, r[2] + folga, r[3] + folga)
