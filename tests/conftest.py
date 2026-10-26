"""Configuração comum dos testes."""

import pytest
from hypothesis import settings

from coir.config import dados_dir

# Testes de propriedade determinísticos (mesmos exemplos em toda execução e no CI).
settings.register_profile("coir", derandomize=True, print_blob=True)
settings.load_profile("coir")


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    """Pula os testes ``golden`` quando a pasta de dados reais (COIR_DADOS) não existe."""
    if dados_dir().is_dir():
        return
    pular = pytest.mark.skip(reason="pasta de dados reais (COIR_DADOS) ausente")
    for item in items:
        if "golden" in item.keywords:
            item.add_marker(pular)
