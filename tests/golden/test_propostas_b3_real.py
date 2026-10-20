"""F7b: as propostas de eventos do pacote público da B3 para o banco real (cópia). Evento que
atinge a posição e não foi lançado precisa estar explicado em
COIR_DADOS/gabarito/propostas-explicadas.csv (fora do repositório: posições do usuário).

Proposta nova falha (lançar o evento ou explicar); explicada que sumiu também falha (tirar do csv).
"""

import csv
from pathlib import Path

import pytest

from coir import referencias
from coir.config import dados_dir
from coir.recursos import pasta_data
from coir.servicos import eventos
from coir.servicos.estado import Estado

pytestmark = [pytest.mark.golden, pytest.mark.slow]
EXPLICADAS = dados_dir() / "gabarito" / "propostas-explicadas.csv"


def test_eventos_do_pacote_lancados_ou_explicados(copia: Path) -> None:
    if not EXPLICADAS.is_file():
        pytest.skip("propostas-explicadas.csv ausente na pasta do gabarito")
    [titular] = Estado(copia).titulares()
    pasta = pasta_data()
    lista = eventos.propostas(copia, titular.id, referencias.carregar(pasta),
                              referencias.carregar_fiscais(pasta), pasta)  # fmt: skip
    do_pacote = [p for p in lista if "eventos da B3" in p.movimentacao]
    assert sum(p.situacao == "lançado" for p in do_pacote) >= 27  # os eventos já lançados
    pendentes = {(p.data.isoformat(), p.codigo, p.dados.tipo.value)
                 for p in do_pacote if p.situacao == "pendente"}  # fmt: skip
    explicadas = {(r["data"], r["codigo"], r["tipo"]): r["motivo"]
                  for r in csv.DictReader(EXPLICADAS.open(encoding="utf-8"))}  # fmt: skip
    assert sorted(pendentes - set(explicadas)) == [], "proposta nova: lançar ou explicar"
    assert sorted(set(explicadas) - pendentes) == [], "resolvida: tirar do csv"
    assert all(explicadas.values())
