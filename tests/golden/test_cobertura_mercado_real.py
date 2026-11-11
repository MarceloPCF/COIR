"""Cobertura do pacote de eventos contra o mercado inteiro, desde 2008, com os arquivos públicos já
baixados (COTAHIST da B3, FRE e FCA da CVM) na pasta de ``COIR_CACHE_B3``: nenhum ato do FRE fica
sem evento no pacote e nenhuma mudança de distribuição com a marca de bonificação, desdobramento
ou grupamento da B3 (ou com salto de preço perto de uma fração simples) fica sem evento, a não ser
o que está conferido em ``sem_evento.csv`` com motivo e fonte.

Sem a pasta (ou sem rede para baixar o que falta), o teste é pulado. Falha lista o que falta.
"""

import csv
import io
import os
from datetime import date
from pathlib import Path

import pytest

from tools.referencias import fontes, gerar, historico

pytestmark = [pytest.mark.golden, pytest.mark.slow]
DATA = Path(__file__).resolve().parents[2] / "data"


def _sem_rede(url: str) -> bytes:
    raise pytest.skip.Exception(f"arquivo fora do cache: {url}")


def test_todo_evento_do_mercado_esta_no_pacote_ou_conferido() -> None:
    pasta = os.environ.get("COIR_CACHE_B3", "")
    if not pasta or not Path(pasta).is_dir():
        pytest.skip("COIR_CACHE_B3 ausente")
    companhias = {r["cnpj"]: r["emissor"] for r in gerar.ler_csv(DATA / "emissores_b3.csv")
                  if r["tipo"] == "COMPANHIA"}  # fmt: skip
    for r in gerar.ler_csv(gerar.EMISSORES_ANTIGOS):
        companhias.setdefault(r["cnpj"], r["emissor"])
    pares = historico.pares_de_troca(gerar.ler_csv(DATA / "trocas_ticker.csv"))
    pacote = (DATA / "pacote.json").read_text(encoding="utf-8")
    ate = date.fromisoformat(pacote.split('"data": "')[1][:10])
    eventos, revisao, series, calendario = fontes.historico_eventos(
        companhias, ate, Path(pasta), obter=_sem_rede, pares=pares)  # fmt: skip
    linhas = list(csv.DictReader(io.StringIO((DATA / "eventos_b3.csv").read_text("utf-8"))))
    conferidos = gerar.sem_evento()
    faltam = [r.texto() for r in historico.cobertas(revisao, linhas, conferidos)]
    faltam += [s.texto() for s in historico.saltos(series, calendario, linhas, conferidos, pares)]
    # rodar o histórico de novo não traz evento novo nem divergência (é o mesmo pacote)
    resumo = gerar.Resumo()
    gerar.juntar_eventos(linhas, [*eventos, *gerar.eventos_curados()], resumo,
                         gerar.eventos_excluidos())  # fmt: skip
    faltam += resumo.eventos_novos + resumo.divergencias + resumo.conflitos
    assert faltam == [], "\n".join(faltam)
