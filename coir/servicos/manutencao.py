"""Backups do banco para a tela de configurações: listar, fazer agora e restaurar."""

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from coir.db.sessao import fazer_backup, listar_backups, restaurar_backup
from coir.erros import BancoError


@dataclass(frozen=True)
class Backup:
    caminho: Path
    feito_em: datetime | None  # pelo nome do arquivo (coir-AAAAMMDD-HHMMSS.db)
    tamanho: int


def _quando(arquivo: Path) -> datetime | None:
    """coir-AAAAMMDD-HHMMSS[-NN].db -> data e hora do backup."""
    try:
        return datetime.strptime("-".join(arquivo.stem.split("-")[1:3]), "%Y%m%d-%H%M%S")
    except ValueError:
        return None


def backups(banco: Path) -> list[Backup]:
    return [Backup(b, _quando(b), b.stat().st_size) for b in listar_backups(banco)]


def fazer_agora(banco: Path, agora: datetime) -> Path:
    destino = fazer_backup(banco, agora)
    if destino is None:
        raise BancoError("ainda não há banco para copiar")
    return destino


def restaurar(banco: Path, origem: Path, agora: datetime) -> Path | None:
    return restaurar_backup(banco, origem, agora)
