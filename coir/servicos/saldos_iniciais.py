"""Prejuízo a compensar em 31/12 de quem começa a usar o COIR sem o histórico de notas (F9).

A custódia inicial é um evento ("Custódia inicial", tela Eventos); o prejuízo fica aqui, um por
titular, e entra na apuração como o prejuízo anterior de janeiro seguinte (engine/apuracao)."""

import logging
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from sqlalchemy import select
from sqlalchemy.orm import Session

from coir.db.modelos import PrejuizoInicial
from coir.db.sessao import conectar
from coir.domain.fiscal import SaldoInicial
from coir.erros import ImportacaoError

log = logging.getLogger(__name__)


def do_banco(s: Session, titular: int | None = None) -> list[SaldoInicial]:
    consulta = select(PrejuizoInicial)
    if titular is not None:
        consulta = consulta.where(PrejuizoInicial.titular_id == titular)
    return [SaldoInicial(titular=r.titular_id, ano=r.ano, comum=r.comum, day_trade=r.day_trade,
                         fii=r.fii) for r in s.scalars(consulta)]  # fmt: skip


def ler(banco: Path, titular: int) -> SaldoInicial | None:
    with conectar(banco) as fabrica, fabrica() as s:
        saldos = do_banco(s, titular)
    return saldos[0] if saldos else None


def gravar(banco: Path, saldo: SaldoInicial | None, titular: int, agora: datetime) -> None:
    """Grava o prejuízo em 31/12 (``None`` apaga). Valores positivos, em reais e centavos."""
    if saldo is not None:
        for nome, valor in (("comum", saldo.comum), ("day trade", saldo.day_trade),
                            ("FII", saldo.fii)):  # fmt: skip
            if valor < 0 or valor != valor.quantize(Decimal("0.01")):
                raise ImportacaoError(f"prejuízo {nome}: informe um valor positivo, em reais")
    with conectar(banco) as fabrica, fabrica.begin() as s:
        atual = s.scalar(select(PrejuizoInicial).where(PrejuizoInicial.titular_id == titular))
        if saldo is None:
            if atual is not None:
                s.delete(atual)
        else:
            if atual is None:
                atual = PrejuizoInicial(titular_id=titular)
                s.add(atual)
            atual.ano, atual.comum, atual.day_trade, atual.fii = (
                saldo.ano,
                saldo.comum,
                saldo.day_trade,
                saldo.fii,
            )
            atual.registrado_em = agora
    log.info("prejuízo inicial %s", "apagado" if saldo is None else f"em 31/12/{saldo.ano}")
