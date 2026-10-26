"""Aviso de versão nova do programa (F9): só avisa, a partir do último release do GitHub."""

import json

import pytest

from coir.erros import AtualizacaoError
from coir.servicos import versao

PAGINA = "https://github.com/MarceloPCF/COIR/releases/tag/v3.0.1"


def _github(resposta: object) -> versao.Baixar:
    def obter(url: str) -> bytes:
        assert url == versao.ULTIMO
        return json.dumps(resposta).encode()

    return obter


@pytest.mark.parametrize(
    ("instalada", "publicada", "avisa"),
    [
        ("3.0.0", "v3.0.1", True),
        ("3.0.0", "v3.1.0", True),
        ("3.0.0.dev0", "v3.0.0", True),  # a de desenvolvimento é anterior ao lançamento
        ("3.0.0", "v3.0.0", False),
        ("3.1.0", "v3.0.9", False),
        ("3.0.0", "v2.2.0", False),  # o último release ainda é o da v2
    ],
)
def test_avisa_so_versao_mais_nova(instalada: str, publicada: str, avisa: bool) -> None:
    nova = versao.nova(instalada, _github({"tag_name": publicada, "html_url": PAGINA}))
    assert (nova is not None) is avisa
    if nova is not None:
        assert (nova.versao, nova.pagina) == (publicada[1:], PAGINA)


@pytest.mark.parametrize(
    "resposta",
    [
        {"tag_name": "v9.0.0", "html_url": PAGINA, "draft": True},
        {"tag_name": "v9.0.0", "html_url": PAGINA, "prerelease": True},
        {"tag_name": "referencias-2099.01.01", "html_url": PAGINA},
        {"tag_name": "v9.0.0", "html_url": "https://exemplo.com/falso"},  # link fora do repositório
        [],
    ],
)
def test_resposta_estranha_nao_avisa(resposta: object) -> None:
    assert versao.nova("3.0.0", _github(resposta)) is None


def test_resposta_que_nao_e_json_falha_alto() -> None:
    with pytest.raises(AtualizacaoError):
        versao.nova("3.0.0", lambda _url: b"<html>")
