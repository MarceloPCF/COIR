"""Cópia do banco real (COIR_DADOS) com as planilhas da B3 importadas, para os golden do app e da
conciliação. Roda sempre na CÓPIA (tmp): nunca migra nem grava o banco real."""

import sqlite3
from datetime import datetime
from pathlib import Path

import pytest

from coir.config import dados_dir
from coir.db.sessao import NOME_BANCO
from coir.servicos.b3 import importar_planilhas
from coir.servicos.estado import Estado

REAL = dados_dir() / NOME_BANCO
PLANILHAS_B3 = dados_dir() / "_relatorios-b3"
AGORA = datetime(2026, 10, 17, 12, 0, 0)


@pytest.fixture(scope="session")
def copia(tmp_path_factory: pytest.TempPathFactory) -> Path:
    if not REAL.is_file():
        pytest.skip("coir.db real ausente")
    destino = tmp_path_factory.mktemp("app") / NOME_BANCO
    origem = sqlite3.connect(f"file:{REAL.as_posix()}?mode=ro", uri=True)
    alvo = sqlite3.connect(destino)
    try:
        with alvo:
            origem.backup(alvo)
    finally:
        alvo.close()
        origem.close()
    estado = Estado(destino)
    estado.preparar(AGORA)  # migra a cópia (o banco real fica como está)
    planilhas = sorted(PLANILHAS_B3.glob("*.xlsx")) if PLANILHAS_B3.is_dir() else []
    titulares = estado.titulares()
    if planilhas and len(titulares) == 1:  # as planilhas da pasta são do titular único
        importar_planilhas(destino, planilhas, titulares[0].id, AGORA)
    return destino
