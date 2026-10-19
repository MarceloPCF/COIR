"""Livro de operações do titular no ano, para a tela (com os custos da nota rateados como no
motor: pelo valor de cada operação, IN RFB 1.585/2015 art. 58 §2º)."""

from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path

from sqlalchemy import select
from sqlalchemy.orm import selectinload

from coir.db import repositorio as repo
from coir.db.modelos import Conta, NotaCorretagem
from coir.db.sessao import conectar
from coir.engine.rateio import ratear


@dataclass(frozen=True)
class LinhaOperacao:
    id_nota: int
    data: date
    corretora: str
    conta: str
    nota: str
    ativo: str
    mercado: str
    cv: str
    day_trade: bool
    quantidade: Decimal
    preco: Decimal
    total: Decimal
    custos: Decimal | None  # BM&F: os custos entram no resultado do dia (não por negócio)


def livro(banco: Path, titular: int | None, ano: int) -> list[LinhaOperacao]:
    consulta = (
        select(NotaCorretagem)
        .join(Conta)
        .where(NotaCorretagem.data_pregao.between(date(ano, 1, 1), date(ano, 12, 31)))
        .options(
            selectinload(NotaCorretagem.operacoes),
            selectinload(NotaCorretagem.conta).selectinload(Conta.corretora),
        )
        .order_by(NotaCorretagem.data_pregao, NotaCorretagem.id)
    )
    if titular is not None:
        consulta = consulta.where(Conta.titular_id == titular)
    linhas: list[LinhaOperacao] = []
    with conectar(banco) as fabrica, fabrica() as s:
        for nota in s.scalars(consulta):
            ops = repo.operacoes_gravadas(nota)
            bovespa = nota.segmento == "bovespa"
            custos = (
                ratear(nota.custos_total, [o.total for o in ops]) if bovespa else [None] * len(ops)
            )
            for o, custo in zip(ops, custos, strict=True):
                linhas.append(
                    LinhaOperacao(
                        id_nota=nota.id, data=o.data, corretora=o.corretora, conta=o.conta,
                        nota=o.nota, ativo=o.papel, mercado=o.mercado, cv=o.cv,
                        day_trade=o.categoria == "DAY_TRADE", quantidade=o.quantidade,
                        preco=o.preco, total=o.total, custos=custo,
                    )
                )  # fmt: skip
    return linhas
