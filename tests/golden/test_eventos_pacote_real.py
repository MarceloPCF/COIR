"""Cobertura do pacote público de eventos (data/eventos_b3.csv) contra os eventos já validados do
usuário (banco real, cópia): todo desdobramento, grupamento e bonificação lançado, de companhia
ainda listada, tem no pacote o evento com a data ex até uma semana antes da data lançada (o
gabarito costuma lançar no dia do crédito) e o mesmo fator: a quantidade lançada é a que o fator
do pacote dá sobre a posição na data com sem o próprio evento. Ajuste de série de opção conta pelo
evento do ativo-objeto.

Falha lista o que falta (o pacote não traz, ou traz com outra data ou fator).
"""

import re
from datetime import date, datetime, timedelta
from decimal import ROUND_FLOOR, Decimal
from pathlib import Path

import pytest

from coir import referencias
from coir.domain.fiscal import TipoEvento
from coir.recursos import pasta_data
from coir.servicos import eventos
from coir.servicos.estado import Estado

pytestmark = [pytest.mark.golden, pytest.mark.slow]
TIPOS = {TipoEvento.DESDOBRAMENTO, TipoEvento.GRUPAMENTO, TipoEvento.BONIFICACAO}
ANTES, DEPOIS = timedelta(days=5), timedelta(days=7)  # data lançada x data ex do pacote
RE_DATA_COM = re.compile(r"data com (\d\d/\d\d/\d{4})")


def _piso(x: Decimal) -> Decimal:
    return x.to_integral_value(ROUND_FLOOR)


def _quantidade(tipo: TipoEvento, base: Decimal, fator: str, serie: bool) -> Decimal:
    """O que o evento dá sobre a posição ``base`` (como as propostas da tela 7)."""
    grupamento = tipo is TipoEvento.GRUPAMENTO
    f = Decimal(fator) / (1 if grupamento else 100)
    if serie:
        return abs(_piso(abs(base) * (f if grupamento else 1 + f)) - abs(base))
    return base - _piso(base * f) if grupamento else _piso(base * f)


def test_todo_evento_validado_esta_no_pacote(copia: Path) -> None:
    [titular] = Estado(copia).titulares()
    pasta = pasta_data()
    emissores = {i[0] for i in referencias.instrumentos(pasta).values()}
    validados = [e.dados for e in eventos.listar(copia, titular.id)
                 if e.dados.tipo in TIPOS and e.dados.ticker[:4] in emissores]  # fmt: skip
    assert len(validados) >= 25
    lista = eventos.propostas(copia, titular.id, referencias.carregar(pasta),
                              referencias.carregar_fiscais(pasta), pasta)  # fmt: skip
    do_pacote = [p for p in lista if "eventos da B3" in p.movimentacao]
    por_isin = referencias.tickers_por_isin(pasta)
    linhas = referencias.eventos_b3(pasta)

    def confere(p: eventos.Proposta, v: eventos.DadosEvento) -> bool:
        d = p.dados
        if (d.tipo, d.ticker, d.conta, d.cv) != (v.tipo, v.ticker, v.conta, v.cv):
            return False
        if not d.data - ANTES <= v.data <= d.data + DEPOIS:
            return False
        achada = RE_DATA_COM.search(p.nota)
        assert achada, p.nota
        data_com = datetime.strptime(achada[1], "%d/%m/%Y").date()
        # a posição da proposta já inclui o evento lançado até a data com: tira
        sinal = 1 if v.cv == "C" else -1
        base = p.quantidade - (sinal * Decimal(v.quantidade or 0) if v.data <= data_com else 0)
        serie = d.observacao.startswith("ajuste da série")
        fatores = {r["fator"] for r in linhas
                   if r["tipo"] == v.tipo.value and date.fromisoformat(r["data_com"]) == data_com
                   and por_isin.get(r["isin"], "")[:4] == v.ticker[:4]}  # fmt: skip
        return any(_quantidade(v.tipo, base, f, serie) == Decimal(v.quantidade or 0)
                   for f in fatores)  # fmt: skip

    faltam = []
    for v in validados:
        if not any(confere(p, v) for p in do_pacote):
            perto = [f"{p.dados.data} {p.dados.cv} {p.dados.quantidade} (posição {p.quantidade})"
                     for p in do_pacote
                     if p.dados.ticker == v.ticker and p.dados.tipo is v.tipo]  # fmt: skip
            faltam.append(f"{v.data} {v.ticker} {v.tipo.value} {v.cv} {v.quantidade}"
                          f" (pacote: {', '.join(perto) or 'nada'})")  # fmt: skip
    assert faltam == [], "\n".join(faltam)
