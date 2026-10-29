"""Notas para o suporte (coir.servicos.suporte): dados para ocultar, envio (texto e imagens
tarjadas), .zip e e-mail; o texto de nota como entrada do leitor e o comando
``coir notas anonimizar``. Só dados fictícios."""

import zipfile
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote

import pytest
from typer.testing import CliRunner

from coir.cli import app
from coir.config import VAR_DADOS
from coir.db.sessao import migrar
from coir.erros import AnonimizacaoError
from coir.parsing.anonimizar import DadosPessoais
from coir.parsing.extrair import extrair_pasta, extrair_pdf
from coir.parsing.texto_pdf import PdfError, gravar_texto, ler_texto, paginas
from coir.servicos import suporte
from coir.servicos.cadastros import adicionar_titular
from tests.apoio import cpf_de_teste
from tools import notas_ficticias
from tools.verificar_dados_pessoais import carregar_termos

FIXTURES = Path(__file__).parents[1] / "fixtures" / "notas"
AGORA = datetime(2026, 10, 28, 12, 0, 0)
AGORA_2 = datetime(2026, 10, 28, 12, 0, 1)
CPF_JOANA = cpf_de_teste("271828182")


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


def test_preparar_envio_grava_texto_e_imagens_com_nome_neutro(
    tmp_path: Path, pdf_ficticio: Path
) -> None:
    envio = suporte.preparar_envio(
        tmp_path, tmp_path / "coir.db", [("Nota de Fulano.pdf", pdf_ficticio.read_bytes())], AGORA
    )
    assert envio.pasta == tmp_path / "suporte" / "20261028-120000"
    [nota] = envio.notas
    assert (nota.origem, nota.destino) == ("Nota de Fulano.pdf", envio.pasta / "nota-01.txt")
    assert nota.imagens == (envio.pasta / "nota-01-pagina-1.png",) and nota.imagens[0].is_file()
    assert nota.paginas == 1 and nota.ocultados > 0 and nota.alertas == 0
    assert envio.titular(nota) == "CPF sem dados para ocultar"  # layout conhecido: não precisa
    texto = nota.destino.read_text(encoding="utf-8") if nota.destino else ""
    roteiro = notas_ficticias.carregar()
    assert cpf_de_teste()[:9] not in texto.replace(".", "")
    assert roteiro.titular.upper() not in texto.upper() and f" {roteiro.conta}" not in texto
    assert (tmp_path / suporte.ARQUIVO_MAPA).is_file()  # o mapa fica, fora do envio
    zip_ = suporte.gerar_zip(envio)
    assert zip_ == envio.pasta / "notas-para-suporte-20261028-120000.zip"
    with zipfile.ZipFile(zip_) as z:
        assert sorted(z.namelist()) == ["nota-01-pagina-1.png", "nota-01.txt"]
    url = suporte.endereco_email(zip_, "3.0.2")
    assert url.startswith("mailto:suporte.coir@gmail.com?subject=COIR%203.0.2")
    assert zip_.name in unquote(url) and str(tmp_path) not in unquote(url)
    with pytest.raises(AnonimizacaoError, match="não é PDF"):
        suporte.preparar_envio(tmp_path, tmp_path / "coir.db", [("n.txt", b"x")], AGORA_2)


def test_envio_sem_nenhuma_nota_aceita_nao_gera_zip(tmp_path: Path) -> None:
    envio = suporte.preparar_envio(
        tmp_path, tmp_path / "coir.db", [("ruim.pdf", b"%PDF-quebrado")], AGORA
    )
    assert envio.aceitas == [] and envio.notas[0].recusas
    with pytest.raises(AnonimizacaoError, match="nenhuma nota"):
        suporte.gerar_zip(envio)


def _pdf_desconhecida(pasta: Path) -> Path:
    texto = (
        "CORRETORA EXEMPLO S.A.\n"
        "Nota de corretagem 123   Pregão 05/03/2026\n"
        f"Titular: Joana Pereira dos Santos   CPF {CPF_JOANA[:3]}.{CPF_JOANA[3:6]}."
        f"{CPF_JOANA[6:9]}-{CPF_JOANA[9:]}\n"
        "Endereco: RUA DAS ACACIAS 98765 - SETOR BUENO\n"
        "Conta: 731942-1\n"
        "C ABCD3 100 10,00 1.000,00\n"
    )
    pasta.mkdir()
    return notas_ficticias.pdf(texto, pasta / "desconhecida.pdf")


def test_envio_de_corretora_desconhecida_usa_os_dados_informados(tmp_path: Path) -> None:
    pdf = _pdf_desconhecida(tmp_path / "pdf")
    recebidos = [("nota.pdf", pdf.read_bytes())]
    [recusada] = suporte.preparar_envio(tmp_path, tmp_path / "coir.db", recebidos, AGORA).notas
    assert recusada.destino is None and "não tem dados para ocultar" in recusada.recusas[0]
    suporte.salvar_titular(tmp_path, DadosPessoais(
        cpf=CPF_JOANA, nome="Joana Pereira dos Santos", contas=["7319421"],
        endereco="Rua das Acácias, 98765", bairro="Setor Bueno",
    ))  # fmt: skip
    envio = suporte.preparar_envio(tmp_path, tmp_path / "coir.db", recebidos, AGORA_2)
    [nota] = envio.notas
    assert nota.recusas == () and nota.destino is not None and len(nota.imagens) == 1
    assert envio.titular(nota) == "Joana Pereira dos Santos"
    texto = nota.destino.read_text(encoding="utf-8").upper()
    assert not any(t in texto for t in ("JOANA", "ACACIAS", "BUENO", "98765", "731942"))
    assert "C ABCD3 100 10,00 1.000,00" in texto


def test_dados_para_ocultar_juntam_o_banco_e_o_informado(tmp_path: Path) -> None:
    banco = tmp_path / "coir.db"
    migrar(banco, AGORA)
    adicionar_titular(banco, CPF_JOANA, "Joana do Banco")
    [d] = suporte.dados_para_ocultar(tmp_path, banco)
    assert (d.nome, d.cpf, d.contas, d.endereco) == ("Joana do Banco", CPF_JOANA, [], "")
    suporte.salvar_titular(
        tmp_path,
        DadosPessoais(
            cpf=CPF_JOANA, nome="JOANA P SANTOS", contas=["1234 ", ""], endereco="Rua A, 1"
        ),
    )
    [d] = suporte.dados_para_ocultar(tmp_path, banco)
    assert (d.nome, d.contas, d.endereco) == ("JOANA P SANTOS", ["1234"], "Rua A, 1")
    assert (tmp_path / suporte.ARQUIVO_DADOS).is_file()
    with pytest.raises(AnonimizacaoError, match="CPF inválido"):
        suporte.salvar_titular(tmp_path, DadosPessoais(cpf="123", nome="x"))
    with pytest.raises(AnonimizacaoError, match="nome"):
        suporte.salvar_titular(tmp_path, DadosPessoais(cpf=CPF_JOANA))
    (tmp_path / suporte.ARQUIVO_DADOS).write_text("{", encoding="utf-8")
    with pytest.raises(AnonimizacaoError, match="ilegíveis"):
        suporte.carregar_dados(tmp_path)


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
