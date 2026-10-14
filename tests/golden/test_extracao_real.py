"""F2: todas as notas reais (COIR_DADOS) extraídas, conciliadas e iguais ao gabarito ou explicadas.

Pulados no CI e onde a pasta de dados não existe. Critério de saída da F2 (docs/PLANO.md §8).
"""

import time

import pytest

from coir import referencias
from coir.config import dados_dir
from coir.domain.nota import NotaExtraida, OperacaoExtraida
from coir.gabarito import candidato
from coir.gabarito.comparar import Tolerancias, carregar_explicacoes, comparar, explicacao
from coir.gabarito.exportar import ARQUIVO_MAPA, PASTA_PRIVADA, carregar
from coir.gabarito.pseudonimos import MapaPseudonimos
from coir.parsing.extrair import ResultadoExtracao, extrair_pasta
from coir.parsing.normalizar import classes_das_series, operacoes

pytestmark = pytest.mark.golden
NOTAS = dados_dir() / "_notas-corretagens"
GABARITO = dados_dir() / "gabarito"
CONHECIDAS = GABARITO / "divergencias-conhecidas.csv"  # na pasta de dados, nunca no repositório


@pytest.fixture(scope="module")
def extracao() -> tuple[ResultadoExtracao, float]:
    if not NOTAS.is_dir():
        pytest.skip("pasta de notas ausente")
    inicio = time.perf_counter()
    res = extrair_pasta(NOTAS)
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
    assert len(res.notas) == 2033
    assert res.rejeitadas == []
    assert [a for a in res.avisos if "opção flexível" not in a] == []
    assert len(res.avisos) == 1
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
    rel = comparar(carregar(GABARITO), cand, Tolerancias.carregar(), ["negocios", "notas"])
    linhas = [f"{d.nivel} {d.chave} {d.campo}" for d in rel.divergencias]
    linhas += rel.faltando + rel.sobrando
    explicacoes = carregar_explicacoes(CONHECIDAS)
    assert [linha for linha in linhas if explicacao(linha, explicacoes) is None] == []
    assert rel.comparados["notas"] >= 1900
