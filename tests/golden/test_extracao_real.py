"""F2: todas as notas reais (COIR_DADOS) extraídas, conciliadas e iguais ao gabarito ou explicadas.

Pulados no CI e onde a pasta de dados não existe. Critério de saída da F2 (docs/PLANO.md §8).
"""

import time
from pathlib import Path

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
# nota cuja compra (SBFG3) foi liquidada fora: o resumo dos negócios não fecha e a nota não entra
REJEITADAS_CONHECIDAS = {"17825207"}


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
    series = classes_das_series(res.notas)
    return [(n, operacoes(n, refs, series)) for n in res.notas]


def test_todas_as_notas_extraidas_e_conciliadas(extracao: tuple[ResultadoExtracao, float]) -> None:
    res, segundos = extracao
    assert {r.motivo.split(":")[0].removeprefix("nota ") for r in res.rejeitadas} == (
        REJEITADAS_CONHECIDAS
    )
    assert len(res.notas) == 2032
    assert any("opção flexível" in a for a in res.avisos)
    assert segundos < 60


def test_candidato_igual_ao_gabarito_ou_explicado(
    normalizadas: list[tuple[NotaExtraida, list[OperacaoExtraida]]], tmp_path: Path
) -> None:
    conhecidas = GABARITO / "divergencias-conhecidas.csv"
    if not conhecidas.is_file():
        pytest.skip("divergencias-conhecidas.csv ausente na pasta do gabarito")
    mapa = MapaPseudonimos.carregar(GABARITO / PASTA_PRIVADA / ARQUIVO_MAPA)
    cand = candidato.montar(normalizadas, mapa)
    rel = comparar(carregar(GABARITO), cand, Tolerancias.carregar(), ["negocios", "notas"])
    linhas = [f"{d.nivel} {d.chave} {d.campo}" for d in rel.divergencias]
    linhas += rel.faltando + rel.sobrando
    explicacoes = carregar_explicacoes(conhecidas)
    assert [linha for linha in linhas if explicacao(linha, explicacoes) is None] == []
    assert rel.comparados["notas"] >= 1900
