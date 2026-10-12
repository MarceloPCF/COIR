"""Gabarito real do usuário (pasta COIR_DADOS). Pulados no CI e onde a pasta não existe."""

from collections import defaultdict
from decimal import Decimal
from pathlib import Path

import pytest

from coir.config import dados_dir
from coir.gabarito.comparar import Tolerancias, comparar
from coir.gabarito.exportar import ARQUIVO_MANIFESTO, ARQUIVOS, carregar, exportar
from coir.gabarito.extrair import ABAS, extrair
from coir.gabarito.leitor import Linhas, ler_abas
from coir.gabarito.modelo import Categoria, Gabarito
from coir.gabarito.pseudonimos import MapaPseudonimos
from tools.verificar_dados_pessoais import carregar_termos, verificar

pytestmark = pytest.mark.golden
PLANILHA = dados_dir() / "_gabarito" / "gabarito.xlsb"


@pytest.fixture(scope="module")
def abas() -> dict[str, Linhas]:
    if not PLANILHA.is_file():
        pytest.skip("planilha-gabarito ausente")
    return ler_abas(PLANILHA, ABAS)


@pytest.fixture(scope="module")
def extraido(abas: dict[str, Linhas]) -> tuple[Gabarito, MapaPseudonimos]:
    mapa = MapaPseudonimos()
    return extrair(abas, mapa), mapa


def test_contagens(extraido: tuple[Gabarito, MapaPseudonimos]) -> None:
    g, mapa = extraido
    ops = defaultdict(int)
    for o in g.operacoes:
        ops[o.categoria] += 1
    assert ops == {Categoria.COMUM: 3081, Categoria.DAY_TRADE: 1247}
    assert (len(g.eventos), len(g.eventos_cadastro)) == (28, 29)
    assert len([p for p in g.posicoes if p.categoria is Categoria.COMUM]) == 3109
    assert len([p for p in g.posicoes if p.categoria is Categoria.DAY_TRADE]) == 1247
    assert {o.data.year for o in g.operacoes} >= set(range(2012, 2027))
    assert min(o.data.year for o in g.operacoes) == 2008
    assert len(mapa.titulares) == 1


def test_resultado_das_operacoes_bate_com_apuracao_mensal(
    extraido: tuple[Gabarito, MapaPseudonimos],
) -> None:
    """Consistência interna da planilha: soma do resultado das operações na competência
    (Mês/Ano da aba *_Op, não a data do pregão) = Resultado da aba *_Cont, ±R$ 1,00 por mês.

    Regra observada na planilha (a validar na F4 com o especialista fiscal): na apuração comum,
    as vendas marcadas "Isento = Sim" ficam fora; no day trade, entram todas.
    """
    g, _ = extraido
    soma: dict[tuple[Categoria, int, int], Decimal] = defaultdict(Decimal)
    for p in g.posicoes:
        if p.resultado and not (p.categoria is Categoria.COMUM and p.isento):
            soma[(p.categoria, p.ano_competencia, p.mes_competencia)] += p.resultado
    divergentes = [
        (
            m.categoria.value,
            m.ano,
            m.mes,
            m.resultado,
            soma.get((m.categoria, m.ano, m.mes), Decimal(0)),
        )
        for m in g.apuracao
        if abs(m.resultado - soma.get((m.categoria, m.ano, m.mes), Decimal(0))) > Decimal("1.00")
    ]
    assert divergentes == []


def test_exportacao_pseudonimizada_e_estavel(
    extraido: tuple[Gabarito, MapaPseudonimos], tmp_path: Path
) -> None:
    g, mapa = extraido
    m1 = exportar(g, mapa, PLANILHA, tmp_path / "a")
    m2 = exportar(g, mapa, PLANILHA, tmp_path / "b")
    assert m1["arquivos"] == m2["arquivos"]
    publicos = [tmp_path / "a" / nome for nome, _ in ARQUIVOS.values()]
    publicos.append(tmp_path / "a" / ARQUIVO_MANIFESTO)
    # contas curtas aparecem por coincidência dentro de valores; elas já são verificadas por
    # igualdade exata nas colunas conta/titular (verificar_pseudonimizacao, chamado no exportar)
    termos = carregar_termos() + list(mapa.titulares)
    assert verificar(publicos, limite_kb=100_000, termos=termos) == []
    recarregado = carregar(tmp_path / "a")
    assert recarregado == g
    assert comparar(g, recarregado, Tolerancias.carregar()).ok
