"""Tipos de coluna: Decimal gravado como TEXT (regra 1), nunca REAL."""

from decimal import Decimal

from sqlalchemy import Dialect, Text
from sqlalchemy.types import TypeDecorator

from coir.erros import BancoError


class DecimalTexto(TypeDecorator[Decimal]):
    """Decimal <-> TEXT exato (``str(Decimal)``). Float é recusado: perderia precisão.

    Em SQL a coluna é texto: ORDER BY, SUM e comparações não são numéricos. Somar e ordenar
    valores sempre em Python (Decimal)."""

    impl = Text
    cache_ok = True

    def process_bind_param(self, value: Decimal | None, dialect: Dialect) -> str | None:
        if value is None:
            return None
        if not isinstance(value, Decimal):
            raise BancoError(f"valor monetário deve ser Decimal, recebido {type(value).__name__}")
        if not value.is_finite():
            raise BancoError(f"valor monetário não finito: {value}")
        return str(value)

    def process_result_value(self, value: str | None, dialect: Dialect) -> Decimal | None:
        return None if value is None else Decimal(value)
