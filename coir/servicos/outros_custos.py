"""Outros custos dedutíveis lançados pelo usuário (ex.: taxas de aluguel de ações - BTC): um valor
por titular, mês e categoria, que reduz o resultado do mês (engine/apuracao)."""

from decimal import Decimal
from pathlib import Path

from sqlalchemy import select
from sqlalchemy.orm import Session

from coir.db.modelos import OutroCusto as Registro
from coir.db.sessao import conectar
from coir.domain.fiscal import Categoria, OutroCusto
from coir.erros import ImportacaoError

CATEGORIAS = (Categoria.COMUM, Categoria.DAY_TRADE)


def do_banco(s: Session, titular: int | None = None) -> list[OutroCusto]:
    consulta = select(Registro).order_by(Registro.ano, Registro.mes, Registro.categoria)
    if titular is not None:
        consulta = consulta.where(Registro.titular_id == titular)
    return [OutroCusto(id=r.id, titular=r.titular_id, ano=r.ano, mes=r.mes,
                       categoria=Categoria(r.categoria), valor=r.valor, descricao=r.descricao)
            for r in s.scalars(consulta)]  # fmt: skip


def listar(banco: Path, titular: int) -> list[OutroCusto]:
    with conectar(banco) as fabrica, fabrica() as s:
        return do_banco(s, titular)


def gravar(banco: Path, titular: int, ano: int, mes: int, categoria: Categoria, valor: Decimal,
           descricao: str = "") -> None:  # fmt: skip
    """Grava o valor do mês e categoria (zero apaga o lançamento)."""
    if categoria not in CATEGORIAS or not 1 <= mes <= 12:
        raise ImportacaoError(f"lançamento inválido: {mes:02d}/{ano} {categoria}")
    if valor < 0 or valor != valor.quantize(Decimal("0.01")):
        raise ImportacaoError("o valor tem de ser positivo, em reais e centavos")
    with conectar(banco) as fabrica, fabrica.begin() as s:
        registro = s.scalar(select(Registro).where(
            Registro.titular_id == titular, Registro.ano == ano, Registro.mes == mes,
            Registro.categoria == categoria.value))  # fmt: skip
        if valor == 0:
            if registro is not None:
                s.delete(registro)
            return
        if registro is None:
            registro = Registro(titular_id=titular, ano=ano, mes=mes, categoria=categoria.value)
            s.add(registro)
        registro.valor = valor
        registro.descricao = descricao.strip()
