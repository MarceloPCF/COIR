import sys
from pathlib import Path

import pytest

from coir.erros import ConfiguracaoError
from coir.recursos import pasta_data

RAIZ_REPO = Path(__file__).resolve().parents[2]


def test_em_desenvolvimento_usa_data_da_raiz() -> None:
    assert pasta_data() == RAIZ_REPO / "data"


def test_no_executavel_usa_meipass(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    (tmp_path / "data").mkdir()
    monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path), raising=False)
    assert pasta_data() == tmp_path / "data"


def test_pasta_ausente_falha_alto(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path), raising=False)
    with pytest.raises(ConfiguracaoError, match="tabelas de referência"):
        pasta_data()
