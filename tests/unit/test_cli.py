from pathlib import Path

import pytest
from typer.testing import CliRunner

from coir import __version__, cli
from coir.cli import app
from coir.config import VAR_DADOS
from coir.erros import ConfiguracaoError

runner = CliRunner()


def test_versao() -> None:
    resultado = runner.invoke(app, ["versao"])
    assert resultado.exit_code == 0
    assert resultado.output.strip() == f"COIR {__version__}"


def test_config_mostra_pasta(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(VAR_DADOS, str(tmp_path))
    resultado = runner.invoke(app, ["config"])
    assert resultado.exit_code == 0
    assert str(tmp_path) in resultado.output
    assert "[existe]" in resultado.output


def test_config_pasta_inexistente(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(VAR_DADOS, str(tmp_path / "nao-existe"))
    resultado = runner.invoke(app, ["--verbose", "config"])
    assert resultado.exit_code == 0
    assert "[NÃO existe]" in resultado.output


def test_main_converte_erro_de_dominio(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    def falha() -> None:
        raise ConfiguracaoError("pasta ausente")

    monkeypatch.setattr(cli, "app", falha)
    with pytest.raises(SystemExit) as saida:
        cli.main()
    assert saida.value.code == 1
    assert capsys.readouterr().err.strip() == "Erro: pasta ausente"


def test_banco_titular_e_importar_pasta_vazia(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from tests.apoio import cpf_de_teste

    monkeypatch.setenv(VAR_DADOS, str(tmp_path))
    assert runner.invoke(app, ["banco", "migrar"]).exit_code == 0
    assert (tmp_path / "coir.db").is_file()
    cpf = cpf_de_teste()
    r = runner.invoke(app, ["titular", "adicionar", cpf, "--nome", "Titular de Teste"])
    assert r.exit_code == 0 and cpf not in r.output and "id 1" in r.output
    assert "Titular de Teste" in runner.invoke(app, ["titular", "listar"]).output
    (tmp_path / "_notas-corretagens").mkdir()
    r = runner.invoke(app, ["importar"])
    assert r.exit_code == 0, r.output
    assert "Notas gravadas: 0" in r.output
    r = runner.invoke(app, ["titular", "renomear", "1", "--nome", "Outro Nome"])
    assert r.exit_code == 0 and "Titular 1: Outro Nome" in r.output
    (tmp_path / "eventos.csv").write_text(
        "data,corretora,conta,ticker,tipo,cv,quantidade,valor,corretora_destino,conta_destino,"
        "observacao,origem\n",
        encoding="utf-8",
    )
    r = runner.invoke(app, ["eventos", "importar"])
    assert r.exit_code == 0 and "Eventos novos: 0" in r.output, r.output
    r = runner.invoke(app, ["apurar", "2024-05"])
    assert r.exit_code == 0 and "nada apurado" in r.output, r.output
    assert runner.invoke(app, ["apurar", "2024-13"]).exit_code != 0
    r = runner.invoke(app, ["carteira", "2024-05-31"])
    assert r.exit_code == 0 and "PM fiscal por titular" in r.output, r.output
    r = runner.invoke(app, ["carteira", "2024-05-31", "--visao", "corretora"])
    assert r.exit_code == 0 and "não é o PM do imposto" in r.output, r.output
    assert runner.invoke(app, ["carteira", "--visao", "banco"]).exit_code != 0


@pytest.mark.parametrize(
    ("argumentos", "mensagem"),
    [
        (["exportar", "--ano", "2026", "--destino", "{arquivo}"], "não foi possível usar a pasta"),
        (["titular", "listar", "--banco", "{arquivo}"], "não é um banco do COIR"),
        (["banco", "migrar", "--banco", "{arquivo}"], "não é um banco do COIR"),
        (["notas", "extrair", "--origem", "{nada}", "--destino", "{saida}"],
         "pasta de notas não encontrada"),
        (["eventos", "importar", "{nada}"], "não foi possível ler"),
        (["eventos", "importar", "{binario}"], "não é um CSV em UTF-8"),
        (["titular", "adicionar", "{cpf}", "--nome", " "], "informe o nome do titular"),
    ],
)  # fmt: skip
def test_entradas_erradas_viram_erro_de_dominio(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, argumentos: list[str], mensagem: str
) -> None:
    """Caminho errado, arquivo que não é banco ou CSV e nome vazio: mensagem curta (CoirError,
    que o ``main`` mostra como "Erro: ..."), nunca o traceback do Python (varredura de 11/2026)."""
    from coir.erros import CoirError

    dados = tmp_path / "dados"
    dados.mkdir()
    monkeypatch.setenv(VAR_DADOS, str(dados))
    assert runner.invoke(app, ["banco", "migrar"]).exit_code == 0
    from tests.apoio import cpf_de_teste

    outro = cpf_de_teste("111444777")
    assert runner.invoke(app, ["titular", "adicionar", outro, "--nome", "T"]).exit_code == 0
    arquivo = tmp_path / "texto.csv"
    arquivo.write_text("isto não é banco nem pasta\n", encoding="utf-8")
    binario = tmp_path / "planilha.csv"
    binario.write_bytes(b"\xff\xfe\x00\x81\x9d" * 50)
    caminhos = {"arquivo": arquivo, "nada": tmp_path / "nao-existe", "saida": dados / "saida",
                "binario": binario, "cpf": cpf_de_teste()}  # fmt: skip
    resultado = runner.invoke(app, [a.format(**caminhos) for a in argumentos])
    assert isinstance(resultado.exception, CoirError), resultado.exception
    assert mensagem in str(resultado.exception)
