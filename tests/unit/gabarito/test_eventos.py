"""Eventos do gabarito (pseudonimizados) -> CSV de eventos do COIR com as contas reais."""

import csv
from pathlib import Path

import pytest

from coir.erros import GabaritoError
from coir.gabarito.eventos import exportar_eventos
from coir.gabarito.pseudonimos import MapaPseudonimos

CABECALHO = "categoria,linha,titular,corretora,conta,data,cv,papel,tipo,preco,quantidade,total\n"


def _gabarito(pasta: Path, *linhas: str) -> Path:
    (pasta / "eventos.csv").write_text(CABECALHO + "".join(linhas), encoding="utf-8")
    mapa = MapaPseudonimos()
    mapa.conta("XP", "12345")  # -> XP-1
    mapa.salvar(pasta / "_privado" / "mapa-pseudonimos.json")
    return pasta


def test_exporta_com_a_conta_real_e_o_tipo_do_coir(tmp_path: Path) -> None:
    pasta = _gabarito(
        tmp_path,
        "COMUM,16,TITULAR-1,XP,XP-1,2018-12-27,C,AAAA3,Bonificação,5.00,100.00000000,"
        "500.00000000\n",
        "COMUM,17,TITULAR-1,XP,XP-1,2018-04-26,C,BBBB3,Desdobramento,0,200.00000000,\n",
    )
    destino = tmp_path / "dados" / "eventos.csv"
    assert exportar_eventos(pasta, destino) == 2
    linhas = list(csv.DictReader(destino.open(encoding="utf-8")))
    assert [(x["conta"], x["tipo"], x["quantidade"], x["valor"]) for x in linhas] == [
        ("12345", "BONIFICACAO", "100", "500"),
        ("12345", "DESDOBRAMENTO", "200", ""),
    ]
    with pytest.raises(GabaritoError, match="já existe"):
        exportar_eventos(pasta, destino)  # não sobrescreve o que o usuário editou


def test_tipo_sem_traducao_ou_conta_fora_do_mapa_falha(tmp_path: Path) -> None:
    pasta = _gabarito(tmp_path, "COMUM,1,TITULAR-1,XP,XP-1,2018-01-02,C,AAAA3,Cisão,0,1,\n")
    with pytest.raises(GabaritoError, match="sem tradução"):
        exportar_eventos(pasta, tmp_path / "a.csv")
    pasta = _gabarito(tmp_path, "COMUM,1,TITULAR-1,XP,XP-9,2018-01-02,C,AAAA3,Grupamento,0,1,\n")
    with pytest.raises(GabaritoError, match="fora do mapa"):
        exportar_eventos(pasta, tmp_path / "b.csv")


def test_nao_grava_dentro_do_repositorio(tmp_path: Path) -> None:
    pasta = _gabarito(tmp_path, "COMUM,1,TITULAR-1,XP,XP-1,2018-01-02,C,AAAA3,Grupamento,0,1,\n")
    with pytest.raises(GabaritoError, match="fora do repositório"):
        exportar_eventos(pasta, Path(__file__).resolve().parent / "eventos.csv")
