"""Conversões puras das células da planilha (floats do Excel) para tipos do domínio."""

from datetime import date, timedelta
from decimal import ROUND_HALF_EVEN, Decimal

from coir.domain.cpf import cpf_valido, so_digitos
from coir.erros import GabaritoError

CASAS = Decimal("1e-8")
EPOCA_EXCEL = date(1899, 12, 30)
ALIAS_CORRETORA = {"ÁGORA": "AGORA"}


def vazio(valor: object) -> bool:
    return valor is None or (isinstance(valor, str) and not valor.strip())


def para_decimal(valor: object) -> Decimal | None:
    """Float do Excel → Decimal com 8 casas (via ``repr``, sem herdar o ruído binário do float)."""
    if vazio(valor):
        return None
    if isinstance(valor, bool):
        raise GabaritoError(f"Valor booleano onde se esperava número: {valor!r}")
    if isinstance(valor, int | float):
        return Decimal(repr(valor)).quantize(CASAS, rounding=ROUND_HALF_EVEN)
    if isinstance(valor, str):
        try:
            return Decimal(valor.strip().replace(",", ".")).quantize(
                CASAS, rounding=ROUND_HALF_EVEN
            )
        except ArithmeticError as erro:
            raise GabaritoError(f"Texto não numérico: {valor!r}") from erro
    raise GabaritoError(f"Tipo inesperado para número: {type(valor).__name__}")


def exigir_decimal(valor: object, campo: str) -> Decimal:
    numero = para_decimal(valor)
    if numero is None:
        raise GabaritoError(f"Campo obrigatório vazio: {campo}")
    return numero


def data_excel(valor: object) -> date | None:
    """Número de série do Excel (sistema 1900) → data."""
    if vazio(valor):
        return None
    if not isinstance(valor, int | float) or isinstance(valor, bool) or valor < 1:
        raise GabaritoError(f"Data inválida (esperado número de série do Excel): {valor!r}")
    return EPOCA_EXCEL + timedelta(days=int(valor))


def exigir_data(valor: object, campo: str) -> date:
    dia = data_excel(valor)
    if dia is None:
        raise GabaritoError(f"Data obrigatória vazia: {campo}")
    return dia


def corretora(nome: object) -> str:
    """Normaliza o nome da corretora ("XP " → "XP", "Agora" → "AGORA")."""
    if vazio(nome):
        raise GabaritoError("Corretora vazia")
    texto = str(nome).strip().upper()
    return ALIAS_CORRETORA.get(texto, texto)


def texto_inteiro(valor: object) -> str | None:
    """Número inteiro guardado como float (nota, conta) → texto sem ".0"."""
    if vazio(valor):
        return None
    if isinstance(valor, float):
        if not valor.is_integer():
            raise GabaritoError(f"Esperado número inteiro: {valor!r}")
        return str(int(valor))
    return str(valor).strip()


def corrigir_cpf(cpf: str, validos: set[str]) -> tuple[str, bool]:
    """Devolve (cpf_só_dígitos, corrigido?).

    CPF com dígito verificador inválido é tratado como erro de digitação quando difere em no
    máximo 2 posições de exatamente um CPF válido da mesma planilha. As mensagens nunca contêm
    o CPF.
    """
    digitos = so_digitos(cpf)
    if cpf_valido(digitos):
        return digitos, False
    proximos = [
        v
        for v in validos
        if len(v) == len(digitos) and sum(a != b for a, b in zip(v, digitos, strict=True)) <= 2
    ]
    if len(proximos) != 1:
        raise GabaritoError(
            f"CPF com dígito verificador inválido sem correção única ({len(proximos)} candidatos)"
        )
    return proximos[0], True
