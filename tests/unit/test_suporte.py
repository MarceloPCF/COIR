"""Notas para o suporte (coir.servicos.suporte), o texto de nota como entrada do leitor e o
comando ``coir notas anonimizar``. Só dados fictícios."""

import zipfile
from datetime import datetime
from pathlib import Path

import pytest
from typer.testing import CliRunner

from coir.cli import app
from coir.config import VAR_DADOS
from coir.erros import AnonimizacaoError
from coir.parsing.extrair import extrair_pasta, extrair_pdf
from coir.parsing.texto_pdf import PdfError, gravar_texto, ler_texto, paginas
from coir.servicos import suporte
from tests.apoio import cpf_de_teste
from tools import notas_ficticias
from tools.verificar_dados_pessoais import carregar_termos

FIXTURES = Path(__file__).parents[1] / "fixtures" / "notas"
AGORA = datetime(2026, 10, 28, 12, 0, 0)


@pytest.fixture(scope="module")
def pdf_ficticio(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Uma nota fictícia em PDF (layout XP, titular e conta inventados)."""
    nome, texto = notas_ficticias.textos(notas_ficticias.carregar())[0]
    return notas_ficticias.pdf(texto, tmp_path_factory.mktemp("pdf") / f"{nome}.pdf")


# ----------------------------------------------------------------------------- texto da nota
def test_texto_gravado_e_lido_igual_com_fim_de_linha_do_pdf(tmp_path: Path) -> None:
    textos = ["linha 1\r\nlinha 2\r\n", "folha 2\r\n"]
    gravar_texto(tmp_path / "n.txt", textos)
    assert ler_texto(tmp_path / "n.txt") == textos == paginas(tmp_path / "n.txt")
    with pytest.raises(PdfError, match="separador"):
        gravar_texto(tmp_path / "x.txt", ["a\fb"])
    (tmp_path / "vazio.txt").write_text("  \f ", encoding="utf-8")
    with pytest.raises(PdfError, match="vazio"):
        paginas(tmp_path / "vazio.txt")


def test_leitor_aceita_o_texto_no_lugar_do_pdf(tmp_path: Path, pdf_ficticio: Path) -> None:
    gravar_texto(tmp_path / "nota.txt", paginas(pdf_ficticio))
    do_pdf, do_texto = extrair_pdf(pdf_ficticio), extrair_pdf(tmp_path / "nota.txt")
    assert do_pdf.notas and [n.model_dump(exclude={"arquivo"}) for n in do_pdf.notas] == [
        n.model_dump(exclude={"arquivo"}) for n in do_texto.notas
    ]
    assert len(extrair_pasta(tmp_path, ".txt").notas) == len(do_pdf.notas)
    assert extrair_pasta(tmp_path).notas == []  # a pasta só lê .txt quando pedido


# ----------------------------------------------------------------------------- serviço
def test_anonimizar_arquivos_grava_e_reusa_o_mapa(tmp_path: Path) -> None:
    origem = FIXTURES / "xp_2026_08_bovespa.txt"
    mapa = tmp_path / "_privado" / "mapa.json"
    [r] = suporte.anonimizar_arquivos([origem], tmp_path / "saida", mapa)
    assert r.recusas == () and r.destino == tmp_path / "saida" / "xp_2026_08_bovespa.txt"
    primeiro = (tmp_path / "saida" / "xp_2026_08_bovespa.txt").read_bytes()
    segredo = suporte.carregar_mapa(mapa).segredo
    [r2] = suporte.anonimizar_arquivos([origem], tmp_path / "outra", mapa)
    assert r2.destino is not None and r2.destino.read_bytes() == primeiro  # mesmo mapa, mesma cópia
    assert suporte.carregar_mapa(mapa).segredo == segredo


def test_anonimizar_arquivos_recusa_sem_gravar(tmp_path: Path) -> None:
    (tmp_path / "termos-sensiveis.txt").write_text("# comentário\nSALA 201\n", encoding="utf-8")
    sensiveis = suporte.termos_sensiveis(tmp_path)
    assert sensiveis == carregar_termos(tmp_path / "termos-sensiveis.txt")  # a mesma lista
    rs = suporte.anonimizar_arquivos(
        [FIXTURES / "xp_2026_08_bovespa.txt", tmp_path / "nao-existe.pdf"],
        tmp_path / "saida", tmp_path / "mapa.json", sensiveis,
    )  # fmt: skip
    assert [r.destino for r in rs] == [None, None]
    assert rs[0].recusas == ("página 1: termo da lista local de termos sensíveis na cópia",)
    assert "não encontrado" in rs[1].recusas[0]
    assert list((tmp_path / "saida").iterdir()) == []


def test_mapa_ilegivel(tmp_path: Path) -> None:
    (tmp_path / "mapa.json").write_text("{", encoding="utf-8")
    with pytest.raises(AnonimizacaoError, match="ilegível"):
        suporte.carregar_mapa(tmp_path / "mapa.json")


def test_pacote_do_suporte_tem_so_texto_com_nome_neutro(tmp_path: Path, pdf_ficticio: Path) -> None:
    pacote = suporte.pacote_suporte(
        tmp_path, [("Nota de Fulano.pdf", pdf_ficticio.read_bytes())], AGORA
    )
    assert pacote.arquivo == tmp_path / "suporte" / "notas-anonimizadas-20261028-120000.zip"
    assert pacote.arquivo is not None
    assert [(r.origem, r.destino) for r in pacote.resultados] == [
        ("Nota de Fulano.pdf", Path("nota-01.txt"))
    ]
    with zipfile.ZipFile(pacote.arquivo) as z:
        assert z.namelist() == ["nota-01.txt"]
        texto = z.read("nota-01.txt").decode("utf-8")
    roteiro = notas_ficticias.carregar()
    assert cpf_de_teste()[:9] not in texto.replace(".", "")
    assert roteiro.titular.upper() not in texto.upper() and f" {roteiro.conta}" not in texto
    assert (tmp_path / suporte.ARQUIVO_MAPA).is_file()  # o mapa fica, fora do .zip
    with pytest.raises(AnonimizacaoError, match="não é PDF"):
        suporte.pacote_suporte(tmp_path, [("nota.txt", b"x")], AGORA)


def test_pacote_sem_nenhuma_nota_aceita_nao_grava_zip(tmp_path: Path) -> None:
    pacote = suporte.pacote_suporte(tmp_path, [("ruim.pdf", b"%PDF-quebrado")], AGORA)
    assert pacote.arquivo is None and pacote.resultados[0].destino is None
    assert not (tmp_path / "suporte").exists()


# ----------------------------------------------------------------------------- CLI
def test_cli_notas_anonimizar(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(VAR_DADOS, str(tmp_path))
    runner = CliRunner()
    r = runner.invoke(app, ["notas", "anonimizar", str(FIXTURES / "rico_2015_11_bovespa.txt")])
    assert r.exit_code == 0, r.output
    assert "1 cópia(s) gravada(s), 0 recusada(s)." in r.output
    assert (tmp_path / "suporte" / "notas-anonimizadas" / "rico_2015_11_bovespa.txt").is_file()
    (tmp_path / "termos-sensiveis.txt").write_text("rico investimentos\n", encoding="utf-8")
    r = runner.invoke(app, ["notas", "anonimizar", str(FIXTURES / "rico_2015_11_bovespa.txt"),
                            "--saida", str(tmp_path / "x")])  # fmt: skip
    assert r.exit_code == 1
    assert "recusada: rico_2015_11_bovespa.txt" in r.output and "termos sensíveis" in r.output
