"""F4: motor fiscal sobre todas as notas reais (+ eventos de COIR_DADOS/eventos.csv) e os 9 DARFs
pagos de 2020-2021 (critério de saída da F4, PLANO §8): conferem até R$ 1,00 ou têm a diferença
explicada (defeitos conhecidos do gabarito v2, docs/regras-fiscais.md "Decisões da F4").

O banco é temporário (tmp_path), nunca o coir.db do usuário.
"""

import csv
from datetime import date, datetime
from decimal import Decimal

import pytest

from coir import referencias
from coir.config import dados_dir
from coir.db import repositorio as repo
from coir.db.sessao import abrir, migrar
from coir.domain.fiscal import Apuracao
from coir.servicos.apuracao import apurar_banco
from coir.servicos.eventos import importar_eventos, ler_csv
from coir.servicos.importacao import RelatorioImportacao, arquivos_da_pasta, importar_arquivos
from tests.golden.conftest import NOTAS_ANONIMIZADAS, SUFIXO_NOTAS, cpf_da_base

pytestmark = pytest.mark.golden
DADOS = dados_dir()
AGORA = datetime(2026, 10, 15, 12, 0, 0)
# situação de cada DARF pago (confere ou explicado, com o valor do motor e o motivo): fica em
# COIR_DADOS/gabarito/darf-explicados.csv, fora do repositório (histórico financeiro do usuário)
EXPLICADOS = DADOS / "gabarito" / "darf-explicados.csv"


@pytest.fixture(scope="module")
def apuracao(tmp_path_factory: pytest.TempPathFactory) -> Apuracao:
    notas, eventos = NOTAS_ANONIMIZADAS, DADOS / "eventos.csv"
    if not (notas.is_dir() and eventos.is_file()):
        pytest.skip("notas ou eventos.csv ausentes")
    banco = tmp_path_factory.mktemp("banco") / "coir.db"
    migrar(banco, AGORA)
    fabrica = abrir(banco)
    with fabrica.begin() as s:
        repo.adicionar_titular(s, cpf_da_base(), "Titular de Teste")
    refs = referencias.carregar()
    with fabrica() as s:
        rel = importar_arquivos(
            s, arquivos_da_pasta(notas, RelatorioImportacao(), SUFIXO_NOTAS), refs, AGORA
        )
    assert rel.rejeitadas == []
    with fabrica() as s:
        importar_eventos(s, ler_csv(eventos))
    fabrica.kw["bind"].dispose()  # type: ignore[union-attr]
    return apurar_banco(banco, refs, referencias.carregar_fiscais(), date(2026, 10, 15))


def test_darfs_pagos_conferem_ou_estao_explicados(apuracao: Apuracao) -> None:
    if not EXPLICADOS.is_file():
        pytest.skip("darf-explicados.csv ausente na pasta do gabarito")
    gabarito = {
        (int(r["ano"]), int(r["mes"])): Decimal(r["valor"])
        for r in csv.DictReader((DADOS / "gabarito" / "darf.csv").open(encoding="utf-8"))
    }
    situacao = {
        (int(r["ano"]), int(r["mes"])): r for r in csv.DictReader(EXPLICADOS.open(encoding="utf-8"))
    }
    assert {m.titular for m in apuracao.meses} == {1}  # gabarito de um titular só
    motor = {(m.ano, m.mes): m.darf for m in apuracao.meses}
    assert set(situacao) == set(gabarito)  # todo DARF pago tem situação registrada
    for k, r in sorted(situacao.items()):
        if r["situacao"] == "confere":
            assert abs(motor.get(k, Decimal(0)) - gabarito[k]) <= 1, k
        else:  # explicado: o motor diverge e o valor dele está registrado (mudou: nova análise)
            assert r["motivo"] and abs(motor.get(k, Decimal(0)) - gabarito[k]) > 1, k
            assert motor.get(k, Decimal(0)) == Decimal(r["motor"]), k


def test_motor_roda_o_livro_inteiro_sem_inconsistencia(apuracao: Apuracao) -> None:
    assert len(apuracao.movimentos) > 9000
    assert apuracao.meses[0].ano == 2008
