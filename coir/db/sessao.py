"""Abertura, migração e backup do banco SQLite (``%COIR_DADOS%\\coir.db``).

Backup antes de migrar e antes de importar (regra 8): cópia consistente pela API de backup do
SQLite em ``<pasta do banco>/backups/``, mantendo as ``MANTER_BACKUPS`` mais recentes.
"""

import logging
import sqlite3
from collections.abc import Iterator
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any

from alembic import command
from alembic.config import Config
from alembic.runtime.migration import MigrationContext
from alembic.script import ScriptDirectory
from sqlalchemy import Engine, create_engine, event
from sqlalchemy.exc import OperationalError
from sqlalchemy.orm import Session, sessionmaker

from coir.erros import BancoError

log = logging.getLogger(__name__)
PASTA_MIGRACOES = Path(__file__).with_name("migracoes")
NOME_BANCO = "coir.db"
MANTER_BACKUPS = 10


def _ligar_chaves_estrangeiras(conexao: Any, _registro: Any) -> None:
    cursor = conexao.cursor()
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.close()


def criar_engine(banco: Path) -> Engine:
    # sem parâmetros nas mensagens de erro: podem trazer CPF e nome (regra 9)
    engine = create_engine(f"sqlite:///{banco}", hide_parameters=True)
    event.listen(engine, "connect", _ligar_chaves_estrangeiras)
    return engine


def _config_alembic() -> Config:
    config = Config()
    config.set_main_option("script_location", str(PASTA_MIGRACOES))
    return config


def revisao_atual(engine: Engine) -> str | None:
    with engine.connect() as conexao:
        return MigrationContext.configure(conexao).get_current_revision()


def revisao_final() -> str | None:
    return ScriptDirectory.from_config(_config_alembic()).get_current_head()


def migrar(banco: Path, agora: datetime, backup: bool = True) -> None:
    """Leva o banco à última revisão (cria se não existir); backup antes, se houver o que migrar
    (``backup=False`` quando quem chama acabou de copiar o banco)."""
    banco.parent.mkdir(parents=True, exist_ok=True)
    engine = criar_engine(banco)
    try:
        atual = revisao_atual(engine)
        if atual == revisao_final():
            return
        if backup and banco.is_file() and atual is not None:
            fazer_backup(banco, agora)
        config = _config_alembic()
        with engine.begin() as conexao:
            config.attributes["connection"] = conexao
            command.upgrade(config, "head")
        log.info("banco migrado de %s para %s: %s", atual, revisao_final(), banco)
    finally:
        engine.dispose()


def _engine_migrado(banco: Path) -> Engine:
    if not banco.is_file():
        raise BancoError(f"banco não encontrado: {banco} (rode `coir banco migrar`)")
    engine = criar_engine(banco)
    if revisao_atual(engine) != revisao_final():
        engine.dispose()
        raise BancoError(f"banco desatualizado: {banco} (rode `coir banco migrar`)")
    return engine


def abrir(banco: Path) -> sessionmaker[Session]:
    """Fábrica de sessões de um banco já migrado; falha alto se o esquema estiver desatualizado.
    Quem chama libera as conexões; prefira ``conectar``."""
    return sessionmaker(_engine_migrado(banco), expire_on_commit=False)


@contextmanager
def conectar(banco: Path) -> Iterator[sessionmaker[Session]]:
    """Como ``abrir``, liberando as conexões ao sair (no Windows, conexão aberta trava o .db)."""
    engine = _engine_migrado(banco)
    try:
        yield sessionmaker(engine, expire_on_commit=False)
    except OperationalError as erro:  # ex.: "database is locked" (outra janela, a CLI)
        raise BancoError(f"banco em uso ou inacessível, tente de novo ({erro.orig})") from erro
    finally:
        engine.dispose()


def fazer_backup(banco: Path, agora: datetime, manter: int = MANTER_BACKUPS) -> Path | None:
    """Copia o banco para ``backups/coir-AAAAMMDD-HHMMSS.db``; sem banco, não faz nada."""
    if not banco.is_file():
        return None
    pasta = banco.parent / "backups"
    pasta.mkdir(exist_ok=True)
    destino, n = pasta / f"{banco.stem}-{agora:%Y%m%d-%H%M%S}.db", 1
    while destino.exists():  # nunca sobrescreve outro backup do mesmo segundo
        n += 1
        destino = pasta / f"{banco.stem}-{agora:%Y%m%d-%H%M%S}-{n:02d}.db"
    origem = sqlite3.connect(banco)
    copia = sqlite3.connect(destino)
    try:
        with copia:
            origem.backup(copia)
    finally:
        copia.close()
        origem.close()
    antigos = sorted(pasta.glob(f"{banco.stem}-*.db"), key=_ordem)[:-manter]
    for velho in antigos:
        velho.unlink()
    log.info("backup do banco: %s", destino)
    return destino


def _ordem(backup: Path) -> tuple[str, int]:
    """coir-AAAAMMDD-HHMMSS[-NN].db -> (data e hora, sequência no mesmo segundo)."""
    partes = backup.stem.split("-")
    return "-".join(partes[1:3]), int(partes[3]) if len(partes) > 3 and partes[3].isdigit() else 1


def listar_backups(banco: Path) -> list[Path]:
    """Backups do banco, do mais recente ao mais antigo."""
    pasta = banco.parent / "backups"
    if not pasta.is_dir():
        return []
    return sorted(pasta.glob(f"{banco.stem}-*.db"), key=_ordem, reverse=True)


def restaurar_backup(banco: Path, origem: Path, agora: datetime) -> Path | None:
    """Volta o banco a um backup da pasta ``backups/``. Antes, o banco atual vira um backup
    (desfaz a restauração); depois, o restaurado é migrado à última revisão."""
    if origem.resolve() not in {b.resolve() for b in listar_backups(banco)}:
        raise BancoError(f"backup desconhecido: {origem.name}")
    memoria = sqlite3.connect(":memory:")
    try:
        leitura = sqlite3.connect(f"file:{origem.as_posix()}?mode=ro", uri=True)
        try:
            leitura.backup(memoria)  # lido antes de girar os backups (o mais antigo sai)
        finally:
            leitura.close()
        atual = fazer_backup(banco, agora)
        destino = sqlite3.connect(banco)
        try:
            with destino:
                memoria.backup(destino)
        finally:
            destino.close()
    except sqlite3.Error as erro:
        raise BancoError(f"não foi possível restaurar {origem.name}: {erro}") from erro
    finally:
        memoria.close()
    log.info("banco restaurado do backup %s (o anterior ficou em %s)", origem.name,
             atual.name if atual else "-")  # fmt: skip
    migrar(banco, agora, backup=False)  # o banco anterior já está em `atual`
    return atual
