"""Configuração comum dos testes."""

from pathlib import Path

import pytest
from hypothesis import settings

from coir.config import VAR_CONFIG, dados_dir

# Testes de propriedade determinísticos (mesmos exemplos em toda execução e no CI).
settings.register_profile("coir", derandomize=True, print_blob=True)
settings.load_profile("coir")


@pytest.fixture(autouse=True)
def _config_isolada(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """A configuração gravada pelo programa no %APPDATA% da máquina (pasta de dados, pasta das
    notas, tema) não vaza para os testes."""
    monkeypatch.setenv(VAR_CONFIG, str(tmp_path / "config-teste" / "config.json"))


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    """Pula os testes ``golden`` quando a pasta de dados reais (COIR_DADOS) não existe."""
    if dados_dir().is_dir():
        return
    pular = pytest.mark.skip(reason="pasta de dados reais (COIR_DADOS) ausente")
    for item in items:
        if "golden" in item.keywords:
            item.add_marker(pular)
