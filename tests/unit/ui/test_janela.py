"""COIR.exe extraído de um .zip baixado: a marca de arquivo baixado sai das DLLs na abertura."""

import sys
from pathlib import Path

import pytest

from coir.ui.janela import desbloquear_dlls

MARCA = "[ZoneTransfer]\nZoneId=3\n"


@pytest.mark.skipif(sys.platform != "win32", reason="Zone.Identifier é um fluxo NTFS")
def test_desbloquear_dlls_tira_a_marca_so_das_dlls(tmp_path: Path) -> None:
    lib = tmp_path / "_internal" / "pythonnet" / "runtime"
    lib.mkdir(parents=True)
    dll, txt, limpa = lib / "Python.Runtime.dll", tmp_path / "LEIA-ME.txt", tmp_path / "x.dll"
    for arquivo in (dll, txt, limpa):
        arquivo.write_bytes(b"x")
    for arquivo in (dll, txt):
        Path(f"{arquivo}:Zone.Identifier").write_text(MARCA, encoding="ascii")

    assert desbloquear_dlls(tmp_path) == 1
    assert not Path(f"{dll}:Zone.Identifier").exists()
    assert Path(f"{txt}:Zone.Identifier").exists()
    assert dll.read_bytes() == b"x"
    assert desbloquear_dlls(tmp_path) == 0
