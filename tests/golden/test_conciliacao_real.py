"""F7a, critério de saída: a posição calculada = a posição da B3 em cada relatório consolidado
importado (fim de cada ano desde 2020 e o mensal mais recente), ou a diferença está explicada em
COIR_DADOS/gabarito/conciliacao-explicada.csv (fora do repositório: posições do usuário).

Diferença nova falha (explicar ou corrigir); diferença explicada que sumiu também falha (tirar do
csv: o evento foi lançado ou o motor mudou).
"""

import csv
from decimal import Decimal
from pathlib import Path

import pytest

from coir.config import dados_dir
from coir.servicos.estado import Estado

pytestmark = [pytest.mark.golden, pytest.mark.slow]
EXPLICADAS = dados_dir() / "gabarito" / "conciliacao-explicada.csv"


def test_posicao_igual_a_da_b3_ou_explicada(copia: Path) -> None:
    if not EXPLICADAS.is_file():
        pytest.skip("conciliacao-explicada.csv ausente na pasta do gabarito")
    estado = Estado(copia)
    [titular] = estado.titulares()
    estado.titular = titular.id
    conciliacoes = estado.conciliacoes()
    assert {c.data.year for c in conciliacoes} >= set(range(2020, 2026))  # dezembro de cada ano
    achadas = {
        (c.data.isoformat(), x.corretora, x.codigo): x for c in conciliacoes for x in c.divergentes
    }
    explicadas = {
        (r["data"], r["corretora"], r["codigo"]): r
        for r in csv.DictReader(EXPLICADAS.open(encoding="utf-8"))
    }
    assert sorted(set(achadas) - set(explicadas)) == [], "diferença nova: explicar ou corrigir"
    assert sorted(set(explicadas) - set(achadas)) == [], "resolvida: tirar do csv"
    for chave, r in explicadas.items():
        x = achadas[chave]
        assert r["motivo"], chave
        assert (x.b3, x.coir) == (Decimal(r["b3"]), Decimal(r["coir"])), chave
