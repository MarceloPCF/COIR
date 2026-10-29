"""F2: todas as notas reais (COIR_DADOS) extraídas, conciliadas e iguais ao gabarito ou explicadas.

Pulados no CI e onde a pasta de dados não existe. Critério de saída da F2 (docs/PLANO.md §8).
"""

import time

import pytest

from coir import referencias
from coir.config import dados_dir
from coir.domain.nota import NotaExtraida, OperacaoExtraida
from coir.gabarito import candidato
from coir.gabarito.comparar import (
    Tolerancias,
    carregar_explicacoes,
    com_tickers_da_data,
    comparar,
    explicacao,
)
from coir.gabarito.exportar import ARQUIVO_MAPA, PASTA_PRIVADA, carregar
from coir.gabarito.pseudonimos import MapaPseudonimos
from coir.parsing.extrair import ResultadoExtracao, extrair_pasta
from coir.parsing.normalizar import classes_das_series, operacoes
from tests.golden.conftest import NOTAS_ANONIMIZADAS, SUFIXO_NOTAS

pytestmark = pytest.mark.golden
NOTAS = NOTAS_ANONIMIZADAS
GABARITO = dados_dir() / "gabarito"
CONHECIDAS = GABARITO / "divergencias-conhecidas.csv"  # na pasta de dados, nunca no repositório


@pytest.fixture(scope="module")
def extracao() -> tuple[ResultadoExtracao, float]:
    if not NOTAS.is_dir():
        pytest.skip("pasta de notas ausente")
    inicio = time.perf_counter()
    res = extrair_pasta(NOTAS, SUFIXO_NOTAS)
    return res, time.perf_counter() - inicio


@pytest.fixture(scope="module")
def normalizadas(
    extracao: tuple[ResultadoExtracao, float],
) -> list[tuple[NotaExtraida, list[OperacaoExtraida]]]:
    res, _ = extracao
    refs = referencias.carregar()
    series = classes_das_series(res.notas, refs)
    return [(n, operacoes(n, refs, series)) for n in res.notas]


def test_todas_as_notas_extraidas_e_conciliadas(extracao: tuple[ResultadoExtracao, float]) -> None:
    res, segundos = extracao
    assert len(res.notas) == 2034
    assert res.rejeitadas == []
    # o XP_NC_2024_11 baixado de novo em 2026 traz também a página BM&F do mês; a nota de opção
    # flexível da BTG estava só na base antiga (o original não está entre as notas reais)
    esperados = ("opção flexível", "repetida (idêntica)")
    assert [a for a in res.avisos if not any(e in a for e in esperados)] == []
    assert sum("repetida (idêntica)" in a for a in res.avisos) == 1
    assert segundos < 60
    if not CONHECIDAS.is_file():
        pytest.skip("divergencias-conhecidas.csv ausente na pasta do gabarito")
    conhecidas = {(e.nota, e.data) for e in carregar_explicacoes(CONHECIDAS)}
    assert [r for r in res.rejeitadas if (r.nota, r.data) not in conhecidas] == []


def test_candidato_igual_ao_gabarito_ou_explicado(
    normalizadas: list[tuple[NotaExtraida, list[OperacaoExtraida]]],
) -> None:
    if not CONHECIDAS.is_file():
        pytest.skip("divergencias-conhecidas.csv ausente na pasta do gabarito")
    mapa = MapaPseudonimos.carregar(GABARITO / PASTA_PRIVADA / ARQUIVO_MAPA)
    cand = candidato.montar(normalizadas, mapa)
    esperado = com_tickers_da_data(carregar(GABARITO), referencias.carregar())
    rel = comparar(esperado, cand, Tolerancias.carregar(), ["negocios", "notas"])
    linhas = [f"{d.nivel} {d.chave} {d.campo}" for d in rel.divergencias]
    linhas += rel.faltando + rel.sobrando
    explicacoes = carregar_explicacoes(CONHECIDAS)
    assert [linha for linha in linhas if explicacao(linha, explicacoes) is None] == []
    assert rel.comparados["notas"] >= 1900
