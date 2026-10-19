"""A interface nunca grava a configuração real do usuário nos testes."""

from pathlib import Path

import pytest

from coir.config import VAR_CONFIG


@pytest.fixture(autouse=True)
def configuracao_temporaria(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    arquivo = tmp_path / "config.json"
    monkeypatch.setenv(VAR_CONFIG, str(arquivo))
    return arquivo
