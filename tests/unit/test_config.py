from pathlib import Path

import pytest

from coir.config import PASTA_PADRAO, VAR_DADOS, dados_dir
from coir.erros import ConfiguracaoError


def test_variavel_de_ambiente_tem_precedencia(tmp_path: Path) -> None:
    assert dados_dir(ambiente={VAR_DADOS: str(tmp_path)}) == tmp_path


def test_sem_variavel_usa_pasta_padrao_na_home() -> None:
    assert dados_dir(ambiente={}) == Path.home() / PASTA_PADRAO


def test_variavel_vazia_usa_padrao() -> None:
    assert dados_dir(ambiente={VAR_DADOS: "  "}) == Path.home() / PASTA_PADRAO


def test_exigir_pasta_inexistente_falha_alto(tmp_path: Path) -> None:
    inexistente = tmp_path / "nao-existe"
    with pytest.raises(ConfiguracaoError, match=VAR_DADOS):
        dados_dir(exigir=True, ambiente={VAR_DADOS: str(inexistente)})


def test_exigir_pasta_existente(tmp_path: Path) -> None:
    assert dados_dir(exigir=True, ambiente={VAR_DADOS: str(tmp_path)}) == tmp_path


def test_le_os_environ_por_padrao(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(VAR_DADOS, str(tmp_path))
    assert dados_dir() == tmp_path
