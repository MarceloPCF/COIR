"""O site publicado (tools/site/montar_site.py): as páginas de site/ e o que elas usam do
repositório, nos endereços de sempre, com todo link interno resolvendo."""

import re
from pathlib import Path
from urllib.parse import unquote, urlparse

from tools.site import montar_site

# a página da v2 aponta para um favicon que nunca existiu (só o ícone da aba falta)
QUEBRADOS_CONHECIDOS = {"legacy/favicon.svg"}
NA_RAIZ = ("index.html", "novidades.html", "fontes.html", "sitemap.xml", "robots.txt", "CNAME",
           "favicon.svg", ".nojekyll")  # fmt: skip


def _links(pagina: Path, raiz: Path) -> list[str]:
    texto = pagina.read_text(encoding="utf-8", errors="replace")
    saida: list[str] = []
    for ref in re.findall(r'(?:href|src)="([^"]+)"', texto) + re.findall(r"url\(([^)]+)\)", texto):
        u = urlparse(ref.strip("'\""))
        if u.scheme in ("http", "https") and u.netloc.endswith("coir.com.br"):
            alvo = raiz / unquote(u.path.lstrip("/"))
        elif u.scheme or not u.path or ref.startswith(("#", "{")):
            continue
        elif u.path.startswith("/"):
            alvo = raiz / unquote(u.path.lstrip("/"))
        else:
            alvo = pagina.parent / unquote(u.path)
        alvo = alvo.resolve()
        if alvo.is_dir() or u.path.endswith("/"):
            alvo = alvo / "index.html"
        saida.append(alvo.relative_to(raiz.resolve()).as_posix())
    return saida


def test_site_montado_com_os_enderecos_de_sempre_e_sem_link_quebrado(tmp_path: Path) -> None:
    destino = tmp_path / "_site"
    montar_site.montar(destino)
    for nome in NA_RAIZ:
        assert (destino / nome).is_file(), nome
    assert list(destino.glob("google*.html"))
    for loc in re.findall(r"<loc>https://coir\.com\.br/([^<]*)</loc>",
                          (destino / "sitemap.xml").read_text(encoding="utf-8")):  # fmt: skip
        assert (destino / (loc + "index.html" if loc.endswith("/") or not loc else loc)).is_file()
    assert (destino / "docs/manual/index.html").is_file()
    assert (destino / "docs/manual/se-nao-abrir.html").is_file()  # link do zip e do programa
    assert (destino / "legacy/index.html").is_file()
    quebrados = {
        alvo
        for pagina in destino.rglob("*.html")
        for alvo in _links(pagina, destino)
        if not (destino / alvo).is_file()
    }
    assert quebrados == QUEBRADOS_CONHECIDOS
