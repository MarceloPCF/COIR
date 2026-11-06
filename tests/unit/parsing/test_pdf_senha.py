"""PDF de nota protegido por senha (as notas enviadas por e-mail costumam vir assim, com o começo
do CPF do titular como senha): a senha informada ou a do CPF de um titular cadastrado abre o PDF;
sem nenhuma que sirva, a nota é rejeitada dizendo que falta a senha (não "PDF ilegível")."""

from datetime import datetime
from pathlib import Path

import pytest
from typer.testing import CliRunner

from coir import referencias
from coir.cli import app
from coir.config import VAR_DADOS
from coir.erros import ExtracaoError
from coir.parsing.extrair import extrair_pdf
from coir.parsing.texto_pdf import PdfSenhaError, paginas, senhas_do_cpf
from coir.servicos.importacao import importar_lista
from coir.servicos.notas import pagina_png
from tests.apoio import cpf_de_teste
from tools import notas_ficticias

AGORA = datetime(2026, 11, 4, 12, 0)
runner = CliRunner()


@pytest.fixture(scope="module")
def textos() -> list[tuple[str, str]]:
    return notas_ficticias.textos(notas_ficticias.carregar())[:2]


def _pdf(pasta: Path, nome: str, texto: str, senha: str | None) -> Path:
    pasta.mkdir(parents=True, exist_ok=True)
    return notas_ficticias.pdf(texto, pasta / f"{nome}.pdf", senha)


def test_senhas_do_cpf_comeco_e_cpf_inteiro_sem_repetir() -> None:
    cpf = cpf_de_teste()
    esperado = [cpf[:3], cpf[:4], cpf[:5], cpf[:6], cpf]
    assert senhas_do_cpf([cpf, f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}", "123"]) == esperado


def test_pdf_com_senha_abre_com_a_senha_certa(
    tmp_path: Path, textos: list[tuple[str, str]]
) -> None:
    nome, texto = textos[0]
    aberto = paginas(_pdf(tmp_path / "aberto", nome, texto, None))
    protegido = _pdf(tmp_path / "protegido", nome, texto, "123")
    with pytest.raises(PdfSenhaError, match="protegido por senha"):
        paginas(protegido)
    with pytest.raises(PdfSenhaError, match="protegido por senha"):
        paginas(protegido, ["999"])
    assert paginas(protegido, ["999", "123"]) == aberto
    [rejeitada] = extrair_pdf(protegido).rejeitadas
    assert "protegido por senha" in rejeitada.motivo


def test_mensagem_orienta_na_tela_e_na_linha_de_comando(
    tmp_path: Path, textos: list[tuple[str, str]]
) -> None:
    """A mesma mensagem chega à tela Importar e ao terminal: precisa servir aos dois."""
    protegido = _pdf(tmp_path, *textos[0], "123")
    with pytest.raises(PdfSenhaError) as erro:
        paginas(protegido)
    assert "Senha dos PDFs" in str(erro.value) and "coir importar --senha" in str(erro.value)


def test_cli_importar_diz_como_informar_a_senha_e_abre_com_ela(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, textos: list[tuple[str, str]]
) -> None:
    """Sem a senha, ``coir importar`` rejeita a nota e ensina a opção; com ela, a nota entra."""
    monkeypatch.setenv(VAR_DADOS, str(tmp_path))
    pasta = tmp_path / "_notas-corretagens"
    _pdf(pasta, *textos[0], "123")
    assert runner.invoke(app, ["banco", "migrar"]).exit_code == 0

    sem_senha = runner.invoke(app, ["importar"])
    assert sem_senha.exit_code == 1
    assert "coir importar --senha" in sem_senha.output

    com_senha = runner.invoke(app, ["importar", "--senha", "123"])
    assert com_senha.exit_code == 0, com_senha.output
    assert "Notas gravadas: 1" in com_senha.output


def test_pagina_da_nota_protegida_na_tela(tmp_path: Path, textos: list[tuple[str, str]]) -> None:
    protegido = _pdf(tmp_path, *textos[0], "123")
    with pytest.raises(ExtracaoError, match="protegido por senha"):
        pagina_png(protegido, 1, 200)
    assert pagina_png(protegido, 1, 200, ["123"]).startswith(b"\x89PNG")


def test_importacao_com_a_senha_informada_e_depois_pelo_cpf_do_titular(
    tmp_path: Path, textos: list[tuple[str, str]]
) -> None:
    """1ª nota: ainda não há titular, só a senha digitada abre. 2ª nota: o titular criado pela 1ª
    tem o CPF, e o começo dele abre o PDF sem senha informada."""
    senha = cpf_de_teste()[:3]
    banco, refs = tmp_path / "coir.db", referencias.carregar()
    primeira = _pdf(tmp_path / "pdfs", *textos[0], senha)
    segunda = _pdf(tmp_path / "pdfs", *textos[1], senha)

    rel = importar_lista(banco, [primeira], refs, AGORA)
    assert rel.notas_gravadas == 0 and "protegido por senha" in rel.rejeitadas[0].motivo
    rel = importar_lista(banco, [primeira], refs, AGORA, senhas=[senha])
    assert rel.notas_gravadas == 1 and not rel.rejeitadas
    rel = importar_lista(banco, [segunda], refs, AGORA)
    assert rel.notas_gravadas == 1 and not rel.rejeitadas
