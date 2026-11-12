"""Licenças de terceiros do pacote: lista de componentes coerente, verificação de cobertura e
nenhum elemento JavaScript do NiceGUI usado pelo COIR fora do pacote."""

import ast
import re
from collections.abc import Callable
from importlib.metadata import files
from pathlib import Path
from typing import cast

import nicegui
import pytest
from nicegui import ui

from tools import licencas

RAIZ = Path(__file__).resolve().parents[2]
SPEC = RAIZ / "empacotamento" / "coir.spec"


def _elementos_do_spec() -> set[str]:
    for no in ast.parse(SPEC.read_text(encoding="utf-8")).body:
        if isinstance(no, ast.Assign) and getattr(no.targets[0], "id", "") == "ELEMENTOS_JS":
            return set(ast.literal_eval(no.value))
    raise AssertionError("ELEMENTOS_JS não encontrado no coir.spec")


def test_cada_arquivo_de_licenca_existe_e_e_usado() -> None:
    lista = licencas.carregar()
    citados = {a for c in lista.componentes for a in c.arquivos if not a.startswith("@")}
    citados |= set(lista.sem_texto.values())
    na_pasta = {p.name for p in licencas.LICENCAS.glob("*.txt")}
    assert citados == na_pasta
    for c in lista.componentes:
        assert c.caminhos and c.licenca and c.fonte.startswith("http"), c.nome


def test_cobertura_acusa_arquivo_sem_licenca_e_componente_velho() -> None:
    lista = licencas.Lista(
        componentes=(
            licencas.Componente(
                "Vue", "3", "MIT", "https://x", ("vue.txt",), ("_internal/nicegui/static/vue*.js",)
            ),
            licencas.Componente(
                "Velho", "1", "MIT", "https://x", ("x.txt",), ("_internal/nicegui/static/velho.js",)
            ),
        ),
        proprios=("_internal/nicegui/static/nicegui.js",),
        vigiadas=("_internal/nicegui/static/**",),
        sem_texto={},
    )
    arquivos = [
        "_internal/nicegui/static/vue.esm-browser.js",
        "_internal/nicegui/static/nicegui.js",
        "_internal/nicegui/static/novo.js",  # biblioteca nova sem licença na lista
        "_internal/nicegui/elements/x.js",  # de pacote Python (RECORD)
        "_internal/coir/__init__.pyc",
        "_internal/misterio.dll",
        "AVISOS-DE-TERCEIROS.txt",
    ]
    problemas = licencas.verificar_cobertura(arquivos, lista, {"_internal/nicegui/elements/x.js"})
    assert problemas == [
        "arquivo sem dono conhecido: _internal/misterio.dll",
        "arquivo de terceiro sem licença na lista: _internal/nicegui/static/novo.js",
        "componente da lista que não está no pacote: Velho",
    ]


def test_ferramentas_de_desenvolvimento_vem_do_pyproject() -> None:
    dev = licencas.de_desenvolvimento()
    assert {"pyinstaller", "pytest", "ruff", "playwright"} <= dev
    assert "nicegui" not in dev


def test_texto_da_licenca_vem_dos_metadados() -> None:
    from importlib import metadata

    textos = dict(licencas.textos_do_pacote(metadata.distribution("pyxlsb")))
    assert "GNU LESSER GENERAL PUBLIC LICENSE" in textos["COPYING.LESSER"]
    assert "GNU GENERAL PUBLIC LICENSE" in textos["COPYING"]  # a LGPL pede as duas
    pdfium = dict(licencas.textos_do_pacote(metadata.distribution("pypdfium2")))
    # caminho, não só o nome (a pasta da plataforma muda: windows_x64, linux_x64...)
    assert any(k.startswith("licenses/data/") and k.endswith("/BUILD_LICENSES/freetype.txt")
               for k in pdfium)  # fmt: skip
    assert licencas.textos_do_pacote(metadata.distribution("proxy-tools")) == []
    assert "proxy-tools" in licencas.carregar().sem_texto  # por isso o texto curado


def test_coir_so_usa_elementos_js_que_vao_no_pacote() -> None:
    """Elemento do NiceGUI com JavaScript próprio (pasta dist/lib) fora de ELEMENTOS_JS não vai
    no pacote: usá-lo quebraria a tela no .exe e deixaria licenças sem aviso."""
    pasta = Path(nicegui.__file__).parent / "elements"
    com_js = {p.parent.name for p in [*pasta.glob("*/dist"), *pasta.glob("*/lib")]}
    permitidos = _elementos_do_spec()
    assert permitidos <= com_js
    # API privada do NiceGUI: se mudar, o teste falha alto e é só reapontar
    preguicosos: dict[str, tuple[str, str]] = ui._LAZY_IMPORTS  # pyright: ignore[reportPrivateUsage]
    proibidos = {nome for nome, (modulo, _) in preguicosos.items()
                 if modulo.split(".")[-1] in com_js - permitidos}  # fmt: skip
    fora = "|".join(sorted(com_js - permitidos))
    usos = re.compile(r"\bui\.(" + "|".join(sorted(proibidos)) + r")\b|make_sortable|mermaid"
                      r"|nicegui\.elements\.(" + fora + r")\b")  # fmt: skip
    achados = [f"{p.relative_to(RAIZ)}: {m.group(0)}"
               for p in (RAIZ / "coir").rglob("*.py")
               for m in usos.finditer(p.read_text(encoding="utf-8"))]  # fmt: skip
    assert achados == [], "inclua o elemento em ELEMENTOS_JS (coir.spec) e a licença dele"


@pytest.mark.parametrize("destino,fora", [
    ("nicegui/elements/mermaid/dist/index.js", True),
    ("nicegui/elements/aggrid/dist/index.js", False),
    ("nicegui/elements/echart/dist/index.js", False),
    ("nicegui/elements/anywidget/lib/a.js", True),
    ("nicegui/elements/markdown.js", False),
    ("reportlab/fonts/Vera.ttf", True),
    ("nicegui/static/vue.esm-browser.js", False),
    # arquivo que o Windows deixou na pasta instalada (não está no RECORD do NiceGUI)
    ("nicegui/InlinePlacementControl.xbf", True),
    ("nicegui/elements/scene/ContextMenu.xbf", True),
])  # fmt: skip
def test_filtro_do_spec(destino: str, fora: bool) -> None:
    arvore = ast.parse(SPEC.read_text(encoding="utf-8"))
    funcao = next(n for n in arvore.body
                  if isinstance(n, ast.FunctionDef) and n.name == "_fora_do_pacote")  # fmt: skip
    do_nicegui = {f.as_posix() for f in files("nicegui") or []}
    escopo: dict[str, object] = {
        "Path": Path, "ELEMENTOS_JS": _elementos_do_spec(), "DO_NICEGUI": do_nicegui,
    }  # fmt: skip
    exec(compile(ast.Module([funcao], []), str(SPEC), "exec"), escopo)
    fora_do_pacote = cast(Callable[[str], bool], escopo["_fora_do_pacote"])
    assert fora_do_pacote(destino) is fora


def test_padroes_de_caminho() -> None:
    """Os padrões de componentes.toml dependem do full_match do Python 3.13."""
    assert licencas.casa("_internal/nicegui/static/lang/pt-BR.js", ["_internal/nicegui/static/**"])
    assert not licencas.casa("_internal/nicegui/static", ["_internal/nicegui/static/**"])
    assert not licencas.casa("_internal/nicegui/static/lang/pt.js", ["_internal/nicegui/static/*"])
    loader = "_internal/webview/lib/runtimes/{}/native/WebView2Loader.dll"
    assert licencas.casa(loader.format("win-x64"), [loader.format("*")])


def test_dlls_do_python_pelo_nome_exato(tmp_path: Path) -> None:
    (tmp_path / "_ssl.pyd").write_bytes(b"")
    (tmp_path / "sqlite3.dll").write_bytes(b"")
    lista = licencas.expandir(licencas.carregar(), tmp_path)
    [python] = [c for c in lista.componentes if c.nome == "Python"]
    assert "_internal/_ssl.pyd" in python.caminhos and "_internal/sqlite3.dll" in python.caminhos
    assert licencas.DLLS_DO_PYTHON not in python.caminhos
    assert not licencas.casa("_internal/_cffi_backend.cp313-win_amd64.pyd", python.caminhos)


def test_lista_vale_so_para_as_versoes_conferidas() -> None:
    conferido = {"nicegui": "3.18.0", "python": "3.13"}
    assert licencas.verificar_versoes(conferido, {"nicegui": "3.18.0", "python": "3.13"}) == []
    [problema] = licencas.verificar_versoes(conferido, {"nicegui": "3.19.0", "python": "3.13"})
    assert "nicegui 3.18.0" in problema and "3.19.0" in problema
    assert set(licencas.carregar().conferido) == {"python", "nicegui", "pywebview"}


def test_avisos_falha_alto_com_ferramenta_de_dev_e_arquivo_sem_dono(tmp_path: Path) -> None:
    interno = tmp_path / "_internal"
    (interno / "pytest").mkdir(parents=True)
    (interno / "pytest" / "__init__.py").write_text("", encoding="utf-8")  # do RECORD do pytest
    (interno / "coir").mkdir()
    (interno / "coir" / "x.pyc").write_bytes(b"")
    (interno / "estranho.xbf").write_bytes(b"")
    lista = licencas.Lista((), (), (), {}, frozenset({"setuptools"}), {})
    with pytest.raises(licencas.AvisosIncompletos) as erro:
        licencas.avisos(tmp_path, "0.0.0", lista, modulos=lambda _: {"coir", "json"})
    mensagem = str(erro.value)
    assert "ferramenta de desenvolvimento dentro do pacote: pytest" in mensagem
    assert "arquivo sem dono conhecido: _internal/estranho.xbf" in mensagem
    assert "coir/x.pyc" not in mensagem


def test_avisos_traz_o_texto_integral(tmp_path: Path) -> None:
    interno = tmp_path / "_internal" / "pyxlsb"
    interno.mkdir(parents=True)
    (interno / "__init__.py").write_text("", encoding="utf-8")
    lista = licencas.Lista((), (), (), {}, frozenset(), {})
    texto = licencas.avisos(tmp_path, "9.9.9", lista, modulos=lambda _: {"pyxlsb"})
    assert texto.startswith("AVISOS DE TERCEIROS — COIR 9.9.9")
    assert "pyxlsb 1.0.10 — LGPLv3+" in texto
    assert "GNU LESSER GENERAL PUBLIC LICENSE" in texto
    assert "The FreeType Project" in texto


def test_texto_ilegivel_falha_alto(tmp_path: Path) -> None:
    (tmp_path / "ok.txt").write_bytes("Copyright © Fulano".encode("cp1252"))
    assert licencas.ler_texto(tmp_path / "ok.txt") == "Copyright © Fulano"
    with pytest.raises(licencas.AvisosIncompletos):
        licencas.ler_texto(tmp_path / "nao-existe.txt")
