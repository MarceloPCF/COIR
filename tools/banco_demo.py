"""Banco de demonstração para a interface: as notas anonimizadas de tests/fixtures/notas com um
titular fictício (CPF sintético). Serve às capturas de tela e aos testes da UI; nenhum dado real.

Com ``pdfs=True``, cada nota vira também um PDF (o texto da fixture, página a página) em
``<pasta>/_notas-corretagens``, para a tela de notas mostrar o original ao lado.

Uso: uv run python -m tools.banco_demo <pasta>   (cria <pasta>/coir.db e os PDFs)
"""

import sys
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

from coir import referencias
from coir.db import repositorio as repo
from coir.db.sessao import NOME_BANCO, conectar, migrar
from coir.parsing.extrair import ResultadoExtracao, extrair_textos
from coir.servicos.importacao import importar_arquivos
from coir.servicos.notas import PASTA_NOTAS
from tests.apoio import cpf_de_teste

FIXTURES = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "notas"
AGORA = datetime(2026, 10, 16, 12, 0, 0)
NOME = "Titular de Demonstração"


def extrator(cpf: str) -> Callable[[Path], ResultadoExtracao]:
    """Lê o texto da fixture (o próprio arquivo, ou a fixture de mesmo nome se for o PDF da
    demonstração) com o CPF anonimizado trocado pelo de teste."""
    mascara = f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}"

    def extrair(caminho: Path) -> ResultadoExtracao:
        fonte = FIXTURES / f"{caminho.stem}.txt" if caminho.suffix == ".pdf" else caminho
        if caminho.suffix == ".pdf" and not caminho.read_bytes().startswith(b"%PDF"):
            fonte = caminho  # texto enviado com nome .pdf (testes da tela de importação)
        texto = fonte.read_text(encoding="utf-8")
        texto = texto.replace("000.000.000-00", f"{mascara}-{cpf[9:]}").replace(
            "000.000.000.00", f"{mascara}.{cpf[9:]}"
        )
        return extrair_textos(caminho.name, texto.split("\f"))

    return extrair


def notas() -> list[Path]:
    """As fixtures, menos o exercício de opção: a série exercida não está nas outras notas e o
    motor falha alto com histórico incompleto."""
    return [p for p in sorted(FIXTURES.glob("*.txt")) if "exercicio" not in p.stem]


def _pdf(texto: str, destino: Path) -> Path:
    """O texto da fixture num PDF de páginas A4 deitadas, em fonte de largura fixa."""
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen.canvas import Canvas

    largura, altura = landscape(A4)
    tela = Canvas(str(destino), pagesize=(largura, altura))
    for pagina in texto.split(""):
        linhas = pagina.splitlines()
        tamanho = min(7.0, (altura - 48) / max(len(linhas), 1) / 1.15)
        tela.setFont("Courier", tamanho)
        y = altura - 28
        for linha in linhas:
            tela.drawString(28, y, linha)
            y -= tamanho * 1.15
        tela.showPage()
    tela.save()
    return destino


def criar(pasta: Path, pdfs: bool = False) -> Path:
    banco = pasta / NOME_BANCO
    pasta.mkdir(parents=True, exist_ok=True)
    migrar(banco, AGORA)
    cpf = cpf_de_teste()
    arquivos = notas()
    if pdfs:
        destino = pasta / PASTA_NOTAS
        destino.mkdir(exist_ok=True)
        arquivos = [_pdf(a.read_text(encoding="utf-8"), destino / f"{a.stem}.pdf")
                    for a in arquivos]  # fmt: skip
    with conectar(banco) as fabrica:
        with fabrica.begin() as s:
            if repo.titular_por_cpf(s, cpf) is None:
                repo.adicionar_titular(s, cpf, NOME)
        with fabrica() as s:
            importar_arquivos(s, arquivos, referencias.carregar(), AGORA,
                              extrator=extrator(cpf))  # fmt: skip
    return banco


if __name__ == "__main__":
    print(criar(Path(sys.argv[1]), pdfs=True))
