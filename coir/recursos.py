"""Localização das tabelas de referência públicas (``data/``), ver docs/adr/0001-pasta-data.md.

Em desenvolvimento, ``data/`` fica na raiz do repositório. No executável PyInstaller (modo pasta),
o spec copia ``data/`` para a pasta de recursos, exposta em ``sys._MEIPASS``.
"""

import sys
from pathlib import Path

from coir.erros import ConfiguracaoError

PASTA = "data"


def pasta_data() -> Path:
    """Retorna a pasta ``data/`` com as tabelas de referência; falha alto se não existir."""
    empacotado: str | None = getattr(sys, "_MEIPASS", None)
    raiz = Path(empacotado) if empacotado else Path(__file__).resolve().parents[1]
    pasta = raiz / PASTA
    if not pasta.is_dir():
        raise ConfiguracaoError(f"Pasta de tabelas de referência não encontrada: {pasta}")
    return pasta
