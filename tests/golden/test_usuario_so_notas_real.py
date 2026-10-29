"""Quem começa só com as notas (sem eventos lançados nem planilhas da B3): o livro para no primeiro
evento que falta (ex.: exercício com a quantidade já ajustada por bonificação), mas a tela de
eventos continua propondo o que falta, senão o usuário fica sem saída (teste da 3.0.1 com as
notas reais). O banco de desenvolvimento tinha os eventos prontos e nunca passou por aqui."""

import sqlite3
from pathlib import Path

import pytest

from coir import referencias
from coir.recursos import pasta_data
from coir.servicos import eventos
from coir.servicos.estado import Estado

pytestmark = [pytest.mark.golden, pytest.mark.slow]


def test_sem_eventos_as_propostas_continuam_disponiveis(copia: Path) -> None:
    with sqlite3.connect(copia) as banco:
        banco.execute("DELETE FROM evento_corporativo")
    [titular] = Estado(copia).titulares()
    pasta = pasta_data()
    lista = eventos.propostas(copia, titular.id, referencias.carregar(pasta),
                              referencias.carregar_fiscais(pasta), pasta)  # fmt: skip
    assert any(p.situacao == "pendente" for p in lista)
