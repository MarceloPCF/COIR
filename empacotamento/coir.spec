# PyInstaller, modo pasta: COIR.exe (janela) e coir-cli.exe (console) na mesma pasta, com
# _internal/ compartilhado. Gerado por tools/empacotar.py (não rodar o PyInstaller direto).
# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

RAIZ = Path(SPECPATH).parent  # noqa: F821 (definido pelo PyInstaller)
AQUI = Path(SPECPATH)  # noqa: F821

dados = [
    (str(RAIZ / "data"), "data"),  # tabelas de referência (coir.recursos.pasta_data)
    (str(RAIZ / "coir" / "db" / "migracoes"), "coir/db/migracoes"),  # o Alembic lê os .py
    (str(RAIZ / "coir" / "parsing" / "profiles"), "coir/parsing/profiles"),
    (str(RAIZ / "coir" / "ui" / "estaticos"), "coir/ui/estaticos"),
    *collect_data_files("nicegui"),
    *collect_data_files("reportlab"),
]
ocultos = [
    *collect_submodules("coir"),
    *collect_submodules("webview"),
    "alembic.runtime.plugins",
    "sqlalchemy.dialects.sqlite",
]
excluir = ["tkinter", "pytest", "hypothesis", "pyright", "ruff", "tests", "tools"]


def analise(script: str) -> Analysis:  # noqa: F821
    return Analysis(  # noqa: F821
        [str(AQUI / script)],
        pathex=[str(RAIZ)],
        datas=dados,
        hiddenimports=ocultos,
        excludes=excluir,
        noarchive=False,
    )


janela, cli = analise("coir_janela.py"), analise("coir_cli.py")
icone = str(AQUI / "coir.ico")
exe_janela = EXE(  # noqa: F821
    PYZ(janela.pure),  # noqa: F821
    janela.scripts, [], exclude_binaries=True, name="COIR", icon=icone,
    console=False, upx=False,
)
exe_cli = EXE(  # noqa: F821
    PYZ(cli.pure),  # noqa: F821
    cli.scripts, [], exclude_binaries=True, name="coir-cli", icon=icone,
    console=True, upx=False,
)
COLLECT(  # noqa: F821
    exe_janela, janela.binaries, janela.datas,
    exe_cli, cli.binaries, cli.datas,
    name="COIR", upx=False,
)
