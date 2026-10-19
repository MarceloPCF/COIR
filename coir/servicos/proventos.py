"""Proventos do titular a partir da movimentação da B3 gravada no banco (engine/proventos)."""

from pathlib import Path

from sqlalchemy import select

from coir.db.modelos import MovimentacaoB3
from coir.db.sessao import conectar
from coir.domain.b3 import LinhaMovimentacao
from coir.domain.parametros import ReferenciasFiscais
from coir.domain.proventos import LinhaB3, Provento
from coir.engine.proventos import proventos


def linhas_banco(banco: Path, titular: int | None = None) -> list[LinhaB3]:
    consulta = select(MovimentacaoB3).order_by(MovimentacaoB3.data, MovimentacaoB3.id)
    if titular is not None:
        consulta = consulta.where(MovimentacaoB3.titular_id == titular)
    with conectar(banco) as fabrica, fabrica() as s:
        return [
            LinhaB3(
                id=m.id,
                titular=m.titular_id,
                linha=LinhaMovimentacao(
                    data=m.data,
                    entrada=m.entrada,
                    movimentacao=m.movimentacao,
                    produto=m.produto,
                    codigo=m.codigo,
                    instituicao=m.instituicao,
                    quantidade=m.quantidade,
                    preco=m.preco,
                    valor=m.valor,
                ),
            )
            for m in s.scalars(consulta)
        ]


def proventos_banco(
    banco: Path, fiscais: ReferenciasFiscais, titular: int | None = None
) -> tuple[list[Provento], list[str]]:
    avisos: list[str] = []
    lista = proventos(linhas_banco(banco, titular), fiscais.tipos, fiscais.parametros, avisos)
    return lista, avisos
