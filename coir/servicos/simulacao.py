"""Tela Simulação: ordens digitadas -> motor (``coir.engine.simulacao``), sem gravar nada.

Lê do banco as mesmas entradas da apuração (notas, eventos, outros custos e saldos iniciais), a
posição por conta para a tela oferecer o que vender, e liga cada ticker digitado ao ativo do
banco (ou a um ativo novo, só na memória, numa compra de ativo que o titular nunca teve).
"""

from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path

from sqlalchemy import select

from coir.db.modelos import Conta
from coir.db.sessao import conectar
from coir.domain.fiscal import CV, Apuracao, TipoAtivo, Visao
from coir.domain.parametros import ReferenciasFiscais
from coir.domain.referencias import Referencias
from coir.engine.apuracao import apurar
from coir.engine.calendario import fim_do_mes
from coir.engine.movimentos import carteira
from coir.engine.simulacao import (
    TIPOS_SIMULAVEIS,
    Entradas,
    OrdemSimulada,
    simular,
    ultimo_mes,
)
from coir.engine.simulacao import ResultadoSimulacao as ResultadoSimulacao  # para a tela
from coir.erros import SimulacaoError
from coir.servicos import outros_custos, saldos_iniciais
from coir.servicos.apuracao import Tickers, entradas


@dataclass(frozen=True)
class ContaSimulavel:
    id: int
    corretora: str
    numero: str

    @property
    def rotulo(self) -> str:
        return f"{self.corretora} {self.numero}"


@dataclass(frozen=True)
class PosicaoSimulavel:
    """Posição da conta que a tela oferece para vender."""

    conta: int
    ticker: str
    tipo: TipoAtivo
    quantidade: Decimal
    pm: Decimal  # na conta (visão operacional); o PM fiscal é o do titular


@dataclass(frozen=True)
class Contexto:
    """O que a tela precisa para montar as ordens: contas, posições e a faixa de datas."""

    contas: tuple[ContaSimulavel, ...]
    posicoes: tuple[PosicaoSimulavel, ...]
    ultima_nota: date | None
    ultimo_mes: tuple[int, int]


@dataclass(frozen=True)
class OrdemDigitada:
    numero: int
    ticker: str
    conta: int
    cv: CV
    quantidade: Decimal
    preco: Decimal
    custos: Decimal
    data: date


@dataclass
class Dados:
    """Entradas do motor e ligações do banco, lidas uma vez por simulação."""

    entradas: Entradas
    contas: dict[int, ContaSimulavel]
    ativo: Callable[[str, date], int | None]


def ler(banco: Path, refs: Referencias, fiscais: ReferenciasFiscais, titular: int) -> Dados:
    with conectar(banco) as fabrica, fabrica() as sessao:
        notas, eventos = entradas(sessao, refs, fiscais)
        tickers = Tickers(sessao)
        contas = {
            c.id: ContaSimulavel(c.id, c.corretora.nome, c.numero)
            for c in sessao.scalars(select(Conta).where(Conta.titular_id == titular))
        }
        dados = Entradas(notas=notas, eventos=eventos, outros_custos=outros_custos.do_banco(sessao),
                         saldos_iniciais=saldos_iniciais.do_banco(sessao))  # fmt: skip

    def ativo(ticker: str, dia: date) -> int | None:
        return tickers.ativo(ticker, dia) or tickers.papel(ticker)

    return Dados(dados, contas, ativo)


def contexto(
    banco: Path, refs: Referencias, fiscais: ReferenciasFiscais, titular: int, hoje: date
) -> Contexto:
    """Contas do titular e as posições por conta (ações, units, ETF, BDR, FII, direitos e FIP)
    no fim da última nota ou de hoje, o que vier depois."""
    dados = ler(banco, refs, fiscais, titular)
    ultima = dados.entradas.ultima_nota(titular)
    dia = max(ultima or hoje, hoje)
    posicoes = carteira(dados.entradas.notas, dados.entradas.eventos, fiscais.parametros, dia,
                        Visao.CONTA, isolados=[])  # fmt: skip
    lista = tuple(
        PosicaoSimulavel(p.conta, p.ticker, p.tipo, p.quantidade, p.pm)
        for p in sorted(posicoes, key=lambda p: (p.ticker, p.conta or 0))
        if p.titular == titular and p.conta is not None and p.quantidade > 0
        and p.tipo in TIPOS_SIMULAVEIS
    )  # fmt: skip
    contas = tuple(sorted(dados.contas.values(), key=lambda c: c.rotulo))
    return Contexto(contas, lista, ultima, ultimo_mes(fiscais.calendario))


def ordens_do_motor(
    dados: Dados, fiscais: ReferenciasFiscais, titular: int, digitadas: Sequence[OrdemDigitada]
) -> list[OrdemSimulada]:
    """Liga ticker e conta: ticker que o titular nunca negociou ganha um ativo só na memória
    (número negativo); o tipo vem das tabelas."""
    novos: dict[str, int] = {}
    saida: list[OrdemSimulada] = []
    for d in digitadas:
        ticker = d.ticker.strip().upper()
        if not ticker:
            raise SimulacaoError(f"ordem {d.numero}: falta o código do ativo")
        conta = dados.contas.get(d.conta)
        if conta is None:
            raise SimulacaoError(f"ordem {d.numero}: escolha a conta")
        ativo = dados.ativo(ticker, d.data)
        if ativo is None:
            ativo = novos.setdefault(ticker, -(len(novos) + 1))
        saida.append(OrdemSimulada(
            numero=d.numero, data=d.data, titular=titular, conta=conta.id,
            corretora=conta.corretora, ativo=ativo, ticker=ticker,
            tipo=fiscais.tipos.tipo_ou_sem(ticker), cv=d.cv, quantidade=d.quantidade,
            preco=d.preco, custos=d.custos,
        ))  # fmt: skip
    return saida


@dataclass(frozen=True)
class Regras:
    """Valores dos parâmetros fiscais numa data, para os textos da tela (nada fixo no texto)."""

    aliquota_comum: Decimal
    aliquota_day_trade: Decimal
    aliquota_fii: Decimal
    limite_isencao: Decimal
    darf_minimo: Decimal
    irrf_day_trade: Decimal


def regras(fiscais: ReferenciasFiscais, dia: date) -> Regras:
    p = fiscais.parametros
    return Regras(
        aliquota_comum=p.decimal("aliquota_comum", dia),
        aliquota_day_trade=p.decimal("aliquota_day_trade", dia),
        aliquota_fii=p.decimal("aliquota_fii", dia),
        limite_isencao=p.decimal("limite_isencao_acoes", dia),
        darf_minimo=p.decimal("darf_minimo", dia),
        irrf_day_trade=p.decimal("aliquota_irrf_day_trade", dia),
    )


def horizonte(cenarios: Mapping[str, Sequence[OrdemDigitada]]) -> date:
    """Fim do mês da última ordem de qualquer cenário: todos são comparados até ele."""
    datas = [o.data for ordens in cenarios.values() for o in ordens]
    if not datas:
        raise SimulacaoError("nenhuma ordem para simular")
    ultima = max(datas)
    return fim_do_mes(ultima.year, ultima.month)


def simular_banco(
    banco: Path, refs: Referencias, fiscais: ReferenciasFiscais, titular: int,
    cenarios: Mapping[str, Sequence[OrdemDigitada]],
    base: Callable[[date], Apuracao] | None = None,
) -> dict[str, ResultadoSimulacao]:  # fmt: skip
    """Cada cenário (nome -> ordens) contra a mesma base, até o mesmo mês. `base`: a apuração
    sem as ordens até uma data (a tela passa a do cache); None calcula aqui. Todos começam no
    mês da primeira ordem de qualquer cenário: os totais cobrem o mesmo período."""
    ate = horizonte(cenarios)
    desde = min(o.data for ordens in cenarios.values() for o in ordens)
    dados = ler(banco, refs, fiscais, titular)
    e = dados.entradas
    apurada = base(ate) if base is not None else apurar(
        e.notas, e.eventos, fiscais.parametros, fiscais.calendario, ate, e.outros_custos,
        e.saldos_iniciais)  # fmt: skip
    saida: dict[str, ResultadoSimulacao] = {}
    for nome, digitadas in cenarios.items():
        if not digitadas:
            continue
        ordens = ordens_do_motor(dados, fiscais, titular, digitadas)
        try:
            saida[nome] = simular(dados.entradas, ordens, fiscais.parametros, fiscais.calendario,
                                  ate, apurada, desde)  # fmt: skip
        except SimulacaoError as erro:
            raise SimulacaoError(f"{nome}: {erro}") from erro
    return saida
