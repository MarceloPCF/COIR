"""Aviso de versão nova do programa nas Releases do GitHub.

Só avisa, com o link da página do release: nunca baixa nem instala nada sozinho (o usuário troca a
pasta do programa; os dados ficam em COIR_DADOS, ver docs/manual/atualizacao.md). Consulta só o
último release do repositório oficial; os pacotes de referências saem marcados como "não é o
último", então o último é sempre o do programa.
"""

import json
import re
from dataclasses import dataclass

from coir.config import VERSAO
from coir.erros import AtualizacaoError
from coir.servicos.atualizacao import REPOSITORIO, Baixar, baixar

ULTIMO = f"https://api.github.com/repos/{REPOSITORIO}/releases/latest"
PAGINAS = f"https://github.com/{REPOSITORIO}/releases/"
RE_VERSAO = re.compile(r"^v?(\d+)\.(\d+)\.(\d+)(.*)$")


@dataclass(frozen=True)
class VersaoRemota:
    versao: str  # "3.0.1"
    pagina: str  # página do release (notas da versão e o .zip)


def chave(versao: str) -> tuple[int, int, int, int]:
    """Ordem das versões: 3.0.0.dev0 < 3.0.0 < 3.0.1 (sufixo = versão de desenvolvimento)."""
    m = RE_VERSAO.match(versao.strip())
    if m is None:
        raise AtualizacaoError(f"versão fora do padrão: {versao!r}")
    return int(m[1]), int(m[2]), int(m[3]), 0 if m[4] else 1


def ultima(obter: Baixar = baixar) -> VersaoRemota | None:
    """O último release publicado do programa (None se a resposta não tiver um)."""
    try:
        dados = json.loads(obter(ULTIMO))
    except json.JSONDecodeError as erro:
        raise AtualizacaoError("resposta inesperada da área de downloads") from erro
    if not isinstance(dados, dict) or dados.get("draft") or dados.get("prerelease"):
        return None
    tag, pagina = str(dados.get("tag_name", "")), str(dados.get("html_url", ""))
    if not RE_VERSAO.match(tag) or not pagina.startswith(PAGINAS):
        return None
    return VersaoRemota(tag.removeprefix("v"), pagina)


def nova(atual: str = VERSAO, obter: Baixar = baixar) -> VersaoRemota | None:
    """A versão publicada, se for mais nova que a instalada."""
    remota = ultima(obter)
    return remota if remota is not None and chave(remota.versao) > chave(atual) else None
