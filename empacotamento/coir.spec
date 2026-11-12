# PyInstaller, modo pasta: COIR.exe (janela) e coir-cli.exe (console) na mesma pasta, com
# _internal/ compartilhado. Gerado por tools/empacotar.py (não rodar o PyInstaller direto).
# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

RAIZ = Path(SPECPATH).parent  # noqa: F821 (definido pelo PyInstaller)
AQUI = Path(SPECPATH)  # noqa: F821

# Componentes JavaScript do NiceGUI que o COIR usa (ui.aggrid e ui.echart). Os outros (Mermaid,
# Plotly, three.js, Leaflet, CodeMirror...) ficam fora do pacote: o NiceGUI só os carrega quando a
# tela usa o elemento, e cada um traria dezenas de bibliotecas a declarar em AVISOS-DE-TERCEIROS.
# Usar outro elemento desses pede incluí-lo aqui e a licença em empacotamento/licencas
# (tests/unit/test_empacotamento.py confere).
ELEMENTOS_JS = {"aggrid", "echart"}

dados = [
    (str(RAIZ / "data"), "data"),  # tabelas de referência (coir.recursos.pasta_data)
    (str(RAIZ / "coir" / "db" / "migracoes"), "coir/db/migracoes"),  # o Alembic lê os .py
    (str(RAIZ / "coir" / "parsing" / "profiles"), "coir/parsing/profiles"),
    (str(RAIZ / "coir" / "parsing" / "rotulos.yaml"), "coir/parsing"),  # rótulos das notas
    (str(RAIZ / "coir" / "ui" / "estaticos"), "coir/ui/estaticos"),
    (str(RAIZ / "docs" / "manual" / "telas" / "*.md"), "docs/manual/telas"),  # "Como faço?"
    *collect_data_files("nicegui"),
    # reportlab: nenhum arquivo de dados. O PDF usa Helvetica, que vem nas métricas embutidas;
    # a pasta reportlab/fonts (DarkGarden GPL, Vera e Type1 sem licença no pacote) não é usada.
]
ocultos = [
    *collect_submodules("coir"),
    # o hook do pywebview para o PyInstaller (webview.__pyinstaller) importaria o PyInstaller
    # inteiro para dentro do pacote
    *collect_submodules("webview", filter=lambda nome: "__pyinstaller" not in nome),
    "alembic.runtime.plugins",
    "sqlalchemy.dialects.sqlite",
]
# Ferramentas de desenvolvimento nunca vão no pacote: _pytest entraria pelo anyio (só para rodar
# testes) e setuptools pelo cffi (só para compilar extensões).
excluir = [
    "tkinter", "pytest", "_pytest", "hypothesis", "pyright", "ruff", "tests", "tools",
    "PyInstaller", "setuptools", "_distutils_hack", "pkg_resources",
]  # fmt: skip


def _fora_do_pacote(destino: str) -> bool:
    partes = Path(destino).parts
    if partes[:2] == ("nicegui", "elements") and len(partes) > 3 and partes[3] in ("dist", "lib"):
        return partes[2] not in ELEMENTOS_JS
    return partes[:2] == ("reportlab", "fonts")


def analise(script: str) -> Analysis:  # noqa: F821
    a = Analysis(  # noqa: F821
        [str(AQUI / script)],
        pathex=[str(RAIZ)],
        datas=dados,
        hiddenimports=ocultos,
        excludes=excluir,
        noarchive=False,
    )
    # os hooks (hook-nicegui coleta a pasta inteira) entram depois de `datas`: o filtro vale no fim
    a.datas = [d for d in a.datas if not _fora_do_pacote(d[0])]
    return a


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
