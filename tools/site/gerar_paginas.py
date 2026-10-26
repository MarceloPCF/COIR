# ruff: noqa: E501  (modelo HTML em linhas longas)
"""Páginas HTML do site a partir do Markdown do repositório, no visual do site (marinho e verde):

- docs/manual/**/*.md -> o .html ao lado (README.md vira index.html), com os links .md trocados;
- docs/fontes-fiscais.md -> fontes.html; CHANGELOG.md -> novidades.html (raiz do site).

O GitHub Pages serve os arquivos como estão (.nojekyll), por isso o Markdown precisa virar HTML.
Uso: uv run python tools/site/gerar_paginas.py
"""

import html
import re
from pathlib import Path

import markdown2

RAIZ = Path(__file__).resolve().parents[2]
SITE = "https://marcelopcf.github.io/COIR/"
RE_LINK = re.compile(r'href="([^":#]+?)\.md(#[^"]*)?"')
RE_TITULO = re.compile(r"^# (.+)$", re.M)

MODELO = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{titulo} · COIR</title>
<link rel="icon" href="{raiz}favicon.svg" type="image/svg+xml">
<style>
:root{{--marinho:#102a43;--verde:#2ea043;--azul-claro:#a8d4ff;--fundo:#f4f7fa;--tinta:#243b53;--linha:#d9e2ec}}
*{{box-sizing:border-box}}
body{{margin:0;font:17px/1.6 "Segoe UI",system-ui,sans-serif;color:var(--tinta);background:var(--fundo)}}
.topo{{background:var(--marinho);border-left:16px solid var(--verde)}}
.topo div{{max-width:980px;margin:0 auto;padding:14px 20px;display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap}}
.topo a{{color:var(--azul-claro);text-decoration:none}}
.topo .marca{{color:#fff;font-weight:800;font-size:22px}}
main{{max-width:980px;margin:0 auto;padding:24px 20px 64px}}
h1{{color:var(--marinho);line-height:1.2}}
h2{{color:var(--marinho);margin-top:2em}}
a{{color:#0b5cad}}
img{{max-width:100%;height:auto;border:1px solid var(--linha);border-radius:6px}}
table{{border-collapse:collapse;width:100%;margin:1em 0;background:#fff;font-size:15px;display:block;overflow-x:auto}}
th,td{{border:1px solid var(--linha);padding:6px 10px;text-align:left;vertical-align:top}}
th{{background:#e9eff5}}
code{{background:#e9eff5;padding:1px 5px;border-radius:4px;font-size:.92em}}
pre{{background:#fff;border:1px solid var(--linha);padding:12px;overflow-x:auto;border-radius:6px}}
pre code{{background:none;padding:0}}
blockquote{{margin:1em 0;padding:8px 16px;border-left:5px solid var(--verde);background:#e8f5ec}}
footer{{max-width:980px;margin:0 auto;padding:0 20px 40px;font-size:14px;color:#627d98}}
</style>
</head>
<body>
<header class="topo"><div><a class="marca" href="{raiz}index.html">COIR</a>
<nav><a href="{raiz}index.html">Início</a> · <a href="{raiz}docs/manual/index.html">Manual</a> · <a href="{raiz}novidades.html">Novidades</a> · <a href="{raiz}fontes.html">Fontes fiscais</a></nav></div></header>
<main>
{corpo}
</main>
<footer>COIR é gratuito e software livre (AGPL-3.0). Não substitui a orientação de um contador.</footer>
</body>
</html>
"""


def _links(corpo: str, origem: Path) -> str:
    """Links relativos para .md viram .html (README -> index; CHANGELOG e fontes -> raiz)."""

    def trocar(m: re.Match[str]) -> str:
        alvo = (origem.parent / f"{m[1]}.md").resolve()
        ancora = m[2] or ""
        if alvo == RAIZ / "CHANGELOG.md":
            destino = RAIZ / "novidades.html"
        elif alvo == RAIZ / "docs" / "fontes-fiscais.md":
            destino = RAIZ / "fontes.html"
        elif alvo.name == "README.md" and alvo.parent == RAIZ / "docs" / "manual":
            destino = alvo.with_name("index.html")
        elif alvo.is_relative_to(RAIZ / "docs" / "manual"):
            destino = alvo.with_suffix(".html")
        else:  # outro Markdown do repositório: a página dele no GitHub
            rel = alvo.relative_to(RAIZ).as_posix()
            return f'href="https://github.com/MarceloPCF/COIR/blob/main/{rel}{ancora}"'
        return f'href="{Path(_relativo(destino, origem.parent)).as_posix()}{ancora}"'

    return RE_LINK.sub(trocar, corpo)


def _relativo(destino: Path, pasta: Path) -> str:
    import os

    return os.path.relpath(destino, pasta).replace(os.sep, "/")


def pagina(origem: Path, destino: Path) -> Path:
    texto = origem.read_text(encoding="utf-8")
    m = RE_TITULO.search(texto)
    titulo = m[1].strip() if m else origem.stem
    corpo = markdown2.markdown(texto, extras=["tables", "fenced-code-blocks", "header-ids"])
    corpo = _links(corpo, origem)
    raiz = _relativo(RAIZ, destino.parent)
    raiz = "" if raiz == "." else raiz + "/"
    destino.write_text(MODELO.format(titulo=html.escape(titulo), raiz=raiz, corpo=corpo),
                       encoding="utf-8", newline="\n")  # fmt: skip
    return destino


def main() -> None:
    manual = RAIZ / "docs" / "manual"
    feitas = [pagina(md, md.with_name("index.html") if md.name == "README.md"
                     else md.with_suffix(".html")) for md in sorted(manual.rglob("*.md"))]  # fmt: skip
    feitas.append(pagina(RAIZ / "docs" / "fontes-fiscais.md", RAIZ / "fontes.html"))
    feitas.append(pagina(RAIZ / "CHANGELOG.md", RAIZ / "novidades.html"))
    for f in feitas:
        print(f.relative_to(RAIZ).as_posix())


if __name__ == "__main__":
    main()
