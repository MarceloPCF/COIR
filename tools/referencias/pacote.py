"""Pacote de referências para o release ``referencias-AAAA.MM.DD``: zip de ``data/`` com
``manifesto.json`` (formato, data e SHA-256 de cada arquivo). O app confere cada hash antes de usar.
"""

from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path

from tools.referencias.gerar import FORMATO, GeracaoError

EXTENSOES = {".csv", ".json", ".toml", ".md"}
MANIFESTO = "manifesto.json"
ZERO = (1980, 1, 1, 0, 0, 0)  # data fixa dentro do zip: mesmo conteúdo, mesmo zip


def _arquivos(pasta: Path) -> list[Path]:
    return sorted(p for p in pasta.iterdir() if p.is_file() and p.suffix in EXTENSOES
                  and p.name != MANIFESTO)  # fmt: skip


def manifesto(pasta: Path) -> dict[str, object]:
    info = json.loads((pasta / "pacote.json").read_text(encoding="utf-8"))
    if info.get("formato") != FORMATO:
        raise GeracaoError(f"pacote.json com formato {info.get('formato')}, esperado {FORMATO}")
    return {
        "formato": FORMATO,
        "data": info["data"],
        "arquivos": {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in _arquivos(pasta)},
    }


def nome_do_release(pasta: Path) -> str:
    data = str(manifesto(pasta)["data"])
    return "referencias-" + data.replace("-", ".")


def montar(pasta: Path, saida: Path) -> Path:
    """Grava ``<saida>/referencias-AAAA.MM.DD.zip`` e o manifesto ao lado (asset do release)."""
    saida.mkdir(parents=True, exist_ok=True)
    dados = manifesto(pasta)
    destino = saida / f"{nome_do_release(pasta)}.zip"
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        for p in _arquivos(pasta):
            z.writestr(zipfile.ZipInfo(p.name, ZERO), p.read_bytes(), zipfile.ZIP_DEFLATED)
        texto = json.dumps(dados, indent=1, sort_keys=True) + "\n"
        z.writestr(zipfile.ZipInfo(MANIFESTO, ZERO), texto, zipfile.ZIP_DEFLATED)
    (saida / MANIFESTO).write_text(texto, encoding="utf-8")
    return destino
