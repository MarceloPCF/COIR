"""Apuração mês a mês x relatórios mensais de outro sistema de IR (um PDF por ano, de 2020 em
diante, em COIR_DADOS/_relatorios-b3). O motor parte do prejuízo que o outro sistema recebeu em
janeiro do primeiro ano (o da DIRPF do ano anterior): diferenças anteriores a esse marco ficam
para a comparação com o gabarito antigo.

Cada linha (mercado, resultado, prejuízo, imposto, IR fonte) confere até a tolerância de mês ou
tem o motivo em COIR_DADOS/gabarito/externo-mensal-explicadas.csv (fora do repositório).
"""

from pathlib import Path

import pytest

from coir import referencias
from coir.config import dados_dir
from coir.engine.calendario import fim_do_mes
from coir.gabarito.comparar import Tolerancias
from coir.gabarito.externo_mensal import (
    carregar_explicacoes,
    comparar_mensal,
    explicar,
    ler_relatorios,
    saldo_inicial,
)
from coir.servicos.apuracao import apurar_banco

pytestmark = pytest.mark.golden
DADOS = dados_dir()
RELATORIOS = sorted((DADOS / "_relatorios-b3").glob("*Comuns e Day-Trade*.pdf"))
EXPLICADAS = DADOS / "gabarito" / "externo-mensal-explicadas.csv"


@pytest.mark.skipif(not RELATORIOS, reason="relatórios mensais do outro sistema ausentes")
def test_apuracao_mensal_confere_com_o_outro_sistema(copia: Path) -> None:
    meses = ler_relatorios(RELATORIOS)
    pasta = referencias.pasta_efetiva(DADOS)
    refs, fiscais = referencias.carregar(pasta), referencias.carregar_fiscais(pasta)
    ultimo = meses[-1]
    apuracao = apurar_banco(copia, refs, fiscais, fim_do_mes(ultimo.ano, ultimo.mes),
                            [saldo_inicial(meses, 1)])  # fmt: skip
    linhas = explicar(comparar_mensal(meses, apuracao.meses, 1, Tolerancias.carregar().mes),
                      carregar_explicacoes(EXPLICADAS))  # fmt: skip
    abertas = [f"{x.mes:02d}/{x.ano} {x.rotulo} ({x.coluna}): motor {x.motor} x {x.externo}"
               for x in linhas if x.classe == "divergente"]  # fmt: skip
    assert not abertas, "\n".join(abertas)
    assert len({(x.ano, x.mes) for x in linhas}) == len(meses)
