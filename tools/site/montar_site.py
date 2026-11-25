"""Monta o site publicado (coir.com.br) numa pasta, com os mesmos endereços de sempre.

As páginas do site ficam em ``site/`` (index, novidades, fontes, guias, sitemap, robots, CNAME...).
O manual (``docs/manual``: o .md vai no zip do programa, o .html é gerado ao lado), as imagens
das telas (``docs/site/imagens``, geradas pelas capturas), a imagem social e a página da v2
(``legacy/``) continuam onde estão e entram no site no mesmo caminho do repositório. O workflow
``site.yml`` publica a pasta montada no GitHub Pages.

Só biblioteca padrão (roda no CI sem instalar o projeto).
Uso: python tools/site/montar_site.py <pasta de saída>
"""

import shutil
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
PAGINAS = RAIZ / "site"
# (pasta ou arquivo do repositório, padrão dos arquivos): publicados no mesmo caminho
DO_REPOSITORIO = (
    ("docs/manual", "**/*.html"),
    ("docs/manual/imagens", "**/*"),
    ("docs/site/imagens", "*"),
    ("docs", "social-preview*.png"),
    ("legacy", "index.html"),
    ("legacy/docs/imagens", "*"),
)


def arquivos() -> dict[str, Path]:
    """Caminho publicado -> arquivo do repositório."""
    saida = {p.relative_to(PAGINAS).as_posix(): p for p in PAGINAS.rglob("*") if p.is_file()}
    for pasta, padrao in DO_REPOSITORIO:
        for p in sorted((RAIZ / pasta).glob(padrao)):
            if p.is_file():
                publicado = p.relative_to(RAIZ).as_posix()
                if publicado in saida:
                    raise SystemExit(f"{publicado} existe em site/ e em {pasta}")
                saida[publicado] = p
    return saida


def montar(destino: Path) -> int:
    if destino.exists():
        shutil.rmtree(destino)
    for publicado, origem in sorted(arquivos().items()):
        alvo = destino / publicado
        alvo.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(origem, alvo)
    return len(list(destino.rglob("*")))


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    total = montar(Path(sys.argv[1]))
    print(f"site montado em {sys.argv[1]} ({total} itens)")
