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
