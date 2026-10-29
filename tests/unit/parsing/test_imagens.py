"""Imagens das páginas com tarja sobre o que o anonimizador trocou. Só notas fictícias."""

from pathlib import Path

import pypdfium2 as pdfium
import pytest

from coir.erros import AnonimizacaoError
from coir.parsing.anonimizar import MapaAnonimizacao, conferir
from coir.parsing.imagens import DPI, gravar_png, paginas_tarjadas
from coir.parsing.texto_pdf import paginas
from tools import notas_ficticias


@pytest.fixture(scope="module")
def pdf(tmp_path_factory: pytest.TempPathFactory) -> Path:
    nome, texto = notas_ficticias.textos(notas_ficticias.carregar())[0]
    return notas_ficticias.pdf(texto, tmp_path_factory.mktemp("pdf") / f"{nome}.pdf")


def test_tarja_cobre_cada_caractere_trocado(pdf: Path, tmp_path: Path) -> None:
    textos = paginas(pdf)
    c = conferir("nota.pdf", textos, MapaAnonimizacao(segredo="00" * 32))
    assert c.aceita
    [imagem] = paginas_tarjadas(pdf, textos, c.trocados)
    cinza = imagem.convert("L")
    escala = DPI / 72
    documento = pdfium.PdfDocument(pdf)
    try:
        pagina = documento[0]
        textpage = pagina.get_textpage()
        caixa = pagina.get_cropbox()
        assert caixa is not None
        topo = caixa[3]
        trocados = [i for i, b in enumerate(c.trocados[0]) if b]
        assert trocados
        cobertos = 0
        for i in trocados:
            esquerda, baixo, direita, cima = textpage.get_charbox(i, loose=True)
            if direita <= esquerda:
                continue
            x = int((esquerda + direita) / 2 * escala)
            y = int((topo - (baixo + cima) / 2) * escala)
            assert cinza.getpixel((x, y)) == 0, f"caractere {i} sem tarja"
            cobertos += 1
        assert cobertos >= len(trocados) // 2
        # e o resto da página continua legível (a tarja não cobre a página toda)
        pretos = cinza.histogram()[0]
        assert pretos < cinza.width * cinza.height / 10
    finally:
        documento.close()
    gravar_png(imagem, tmp_path / "p.png")
    assert 0 < (tmp_path / "p.png").stat().st_size < 300_000


def test_texto_que_nao_casa_com_o_pdf_nao_gera_imagem(pdf: Path) -> None:
    textos = paginas(pdf)
    with pytest.raises(AnonimizacaoError, match="não casa"):
        paginas_tarjadas(pdf, [textos[0] + "x"], [bytes(len(textos[0]) + 1)])
    with pytest.raises(AnonimizacaoError, match="número de páginas"):
        paginas_tarjadas(pdf, [*textos, "outra"], [b"", b""])
