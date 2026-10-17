"""DARFs do ano para a tela: situação (a pagar, paga, em atraso, acumulada abaixo do mínimo), a
multa de mora (motor, Lei 9.430/1996 art. 61) e o registro do pagamento com o comprovante.

A DARF é a do motor (``MesApurado.darf``); o pagamento é um fato lançado pelo usuário e não
muda a apuração. Juros Selic ficam para a F7.
"""

import calendar
import logging
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from enum import StrEnum
from pathlib import Path

from sqlalchemy import select

from coir.db.modelos import DarfPagamento, Titular
from coir.db.sessao import conectar
from coir.domain.fiscal import MesApurado, MultaMora
from coir.domain.parametros import ReferenciasFiscais
from coir.engine.darf import multa_mora
from coir.erros import ImportacaoError

log = logging.getLogger(__name__)
PASTA_COMPROVANTES = "comprovantes"
LIMITE_COMPROVANTE = 10 * 1024 * 1024
ZERO = Decimal(0)


class Situacao(StrEnum):
    A_PAGAR = "a pagar"
    EM_ATRASO = "em atraso"
    PAGA = "paga"
    ACUMULADA = "acumulada"  # abaixo do mínimo: soma à do mês seguinte (Lei 9.430 art. 68 §1º)
    EM_APURACAO = "em apuração"  # o mês ainda não acabou: o valor é parcial


@dataclass(frozen=True)
class Pagamento:
    data: date
    valor: Decimal
    comprovante: str | None  # relativo à pasta de dados
    observacao: str


@dataclass(frozen=True)
class LinhaDarf:
    ano: int
    mes: int
    codigo: str
    vencimento: date | None
    principal: Decimal
    situacao: Situacao
    multa: MultaMora | None  # paga: até o pagamento; em atraso: estimada até hoje
    pagamento: Pagamento | None

    @property
    def total(self) -> Decimal:
        return self.principal + (self.multa.valor if self.multa else ZERO)

    @property
    def pago_em_atraso(self) -> bool:
        p, v = self.pagamento, self.vencimento
        return p is not None and v is not None and p.data > v

    @property
    def diferenca(self) -> Decimal | None:
        """Pago menos devido (principal + multa): negativo = falta pagar; positivo = a maior."""
        return None if self.pagamento is None else self.pagamento.valor - self.total


def pagamentos(banco: Path, titular: int) -> dict[tuple[int, int], Pagamento]:
    with conectar(banco) as fabrica, fabrica() as s:
        linhas = s.scalars(select(DarfPagamento).where(DarfPagamento.titular_id == titular))
        return {
            (p.ano, p.mes): Pagamento(p.data_pagamento, p.valor_pago, p.comprovante, p.observacao)
            for p in linhas
        }


def darfs(
    meses: list[MesApurado], pagos: dict[tuple[int, int], Pagamento], hoje: date,
    fiscais: ReferenciasFiscais,
) -> list[LinhaDarf]:  # fmt: skip
    """Uma linha por mês com DARF ou com valor acumulado abaixo do mínimo."""
    linhas: list[LinhaDarf] = []
    for m in meses:
        competencia = date(m.ano, m.mes, 1)
        codigo = fiscais.parametros.texto("darf_codigo", competencia)
        pago = pagos.get((m.ano, m.mes))
        if (m.ano, m.mes) >= (hoje.year, hoje.month):
            if m.darf > 0 or m.darf_acumulado > 0:
                linhas.append(LinhaDarf(m.ano, m.mes, codigo, None, m.darf + m.darf_acumulado,
                                        Situacao.EM_APURACAO, None, pago))  # fmt: skip
        elif m.darf > 0 and m.vencimento is not None:
            multa: MultaMora | None = None
            if pago is not None:
                situacao = Situacao.PAGA
                multa = multa_mora(m.darf, competencia, m.vencimento, pago.data,
                                   fiscais.parametros, fiscais.calendario)  # fmt: skip
            elif hoje > m.vencimento:
                situacao = Situacao.EM_ATRASO
                multa = multa_mora(m.darf, competencia, m.vencimento, hoje, fiscais.parametros,
                                   fiscais.calendario)  # fmt: skip
            else:
                situacao = Situacao.A_PAGAR
            linhas.append(LinhaDarf(m.ano, m.mes, codigo, m.vencimento, m.darf, situacao, multa,
                                    pago))  # fmt: skip
        elif m.darf_acumulado > 0:
            linhas.append(LinhaDarf(m.ano, m.mes, codigo, None, m.darf_acumulado,
                                    Situacao.ACUMULADA, None, pago))  # fmt: skip
    return linhas


def registrar_pagamento(
    banco: Path, dados: Path, titular: int, ano: int, mes: int, data_pagamento: date,
    valor: Decimal, agora: datetime, comprovante: tuple[str, bytes] | None = None,
    observacao: str = "",
) -> None:  # fmt: skip
    """Grava (ou substitui) o pagamento da DARF da competência; o comprovante (PDF) vai para
    ``<dados>/comprovantes/<ano>/`` com nome sem CPF."""
    if not 1 <= mes <= 12:
        raise ImportacaoError(f"competência inválida: {mes:02d}/{ano}")
    if valor <= 0:
        raise ImportacaoError("informe o valor pago")
    if data_pagamento > agora.date():
        raise ImportacaoError("a data do pagamento não pode ser futura")
    fim = date(ano, mes, calendar.monthrange(ano, mes)[1])
    if data_pagamento < fim:
        raise ImportacaoError("a data do pagamento é anterior ao fim da competência")
    relativo: Path | None = None
    if comprovante is not None:
        nome, conteudo = comprovante
        if not conteudo.startswith(b"%PDF"):
            raise ImportacaoError(f"o comprovante não é PDF: {Path(nome).name}")
        if len(conteudo) > LIMITE_COMPROVANTE:
            raise ImportacaoError("o comprovante passa de 10 MB")
        relativo = (
            Path(PASTA_COMPROVANTES) / str(ano) / f"darf-{ano}-{mes:02d}-titular{titular}.pdf"
        )
    with conectar(banco) as fabrica, fabrica.begin() as s:
        if s.get(Titular, titular) is None:
            raise ImportacaoError(f"titular {titular} não cadastrado")
        registro = s.scalar(select(DarfPagamento).where(
            DarfPagamento.titular_id == titular, DarfPagamento.ano == ano,
            DarfPagamento.mes == mes))  # fmt: skip
        if registro is None:
            registro = DarfPagamento(titular_id=titular, ano=ano, mes=mes)
            s.add(registro)
        registro.data_pagamento = data_pagamento
        registro.valor_pago = valor
        registro.observacao = observacao.strip()
        registro.registrado_em = agora
        if relativo is not None and comprovante is not None:
            try:  # dentro da transação: se o arquivo falhar, o registro não entra
                (dados / relativo).parent.mkdir(parents=True, exist_ok=True)
                (dados / relativo).write_bytes(comprovante[1])
            except OSError as erro:
                raise ImportacaoError(f"não foi possível guardar o comprovante: {erro}") from erro
            registro.comprovante = relativo.as_posix()
    log.info("pagamento de DARF registrado: titular %d, %02d/%d", titular, mes, ano)


def excluir_pagamento(banco: Path, titular: int, ano: int, mes: int) -> None:
    """Desfaz o registro (o comprovante continua na pasta de dados)."""
    with conectar(banco) as fabrica, fabrica.begin() as s:
        registro = s.scalar(select(DarfPagamento).where(
            DarfPagamento.titular_id == titular, DarfPagamento.ano == ano,
            DarfPagamento.mes == mes))  # fmt: skip
        if registro is not None:
            s.delete(registro)
