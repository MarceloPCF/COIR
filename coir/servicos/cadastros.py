"""Cadastros pela interface: titulares (adicionar e renomear) e a consulta de contas, corretoras
e tickers, que nascem da importação e de data/."""

from dataclasses import dataclass
from datetime import date
from pathlib import Path

from sqlalchemy import func, select

from coir.db import repositorio as repo
from coir.db.modelos import Ativo, Conta, Corretora, NotaCorretagem, TickerHistorico, Titular
from coir.db.sessao import conectar
from coir.domain.cpf import cpf_valido, so_digitos
from coir.erros import ImportacaoError


def adicionar_titular(banco: Path, cpf: str, nome: str) -> int:
    """Cadastra o titular (CPF com dígitos verificadores conferidos) e devolve o id."""
    if not nome.strip():
        raise ImportacaoError("informe o nome do titular")
    if not cpf_valido(cpf):
        raise ImportacaoError("CPF inválido (confira os dígitos)")
    with conectar(banco) as fabrica, fabrica.begin() as s:
        return repo.adicionar_titular(s, so_digitos(cpf), nome).id


def renomear_titular(banco: Path, id_titular: int, nome: str) -> None:
    with conectar(banco) as fabrica, fabrica.begin() as s:
        repo.renomear_titular(s, id_titular, nome)


@dataclass(frozen=True)
class TitularCadastro:
    id: int
    nome: str
    cpf: str  # mascarado: só três dígitos do meio (sem os verificadores)
    contas: tuple[str, ...]  # "XP 0000123"


@dataclass(frozen=True)
class ContaCadastro:
    id: int
    corretora: str
    numero: str
    titular: str
    notas: int
    primeira: date | None
    ultima: date | None


@dataclass(frozen=True)
class TickerCadastro:
    ativo: int
    ticker: str
    inicio: date | None
    fim: date | None
    isin: str


def mascarar_cpf(cpf: str) -> str:
    d = so_digitos(cpf)
    return f"***.***.{d[6:9]}-**" if len(d) == 11 else "***"


def titulares(banco: Path) -> list[TitularCadastro]:
    with conectar(banco) as fabrica, fabrica() as s:
        contas = s.execute(
            select(Conta.titular_id, Corretora.nome, Conta.numero)
            .join(Corretora)
            .order_by(Corretora.nome, Conta.numero)
        ).all()
        return [
            TitularCadastro(
                t.id,
                t.nome,
                mascarar_cpf(t.cpf),
                tuple(f"{c} {n}" for tid, c, n in contas if tid == t.id),
            )
            for t in repo.titulares(s)
        ]


def contas(banco: Path) -> list[ContaCadastro]:
    with conectar(banco) as fabrica, fabrica() as s:
        linhas = s.execute(
            select(
                Conta.id,
                Corretora.nome,
                Conta.numero,
                Titular.nome,
                func.count(NotaCorretagem.id),
                func.min(NotaCorretagem.data_pregao),
                func.max(NotaCorretagem.data_pregao),
            )
            .join(Corretora, Conta.corretora_id == Corretora.id)
            .join(Titular, Conta.titular_id == Titular.id)
            .outerjoin(NotaCorretagem, NotaCorretagem.conta_id == Conta.id)
            .group_by(Conta.id)
            .order_by(Corretora.nome, Conta.numero)
        ).all()
    return [ContaCadastro(*linha) for linha in linhas]


def tickers(banco: Path) -> list[TickerCadastro]:
    with conectar(banco) as fabrica, fabrica() as s:
        linhas = s.execute(
            select(
                TickerHistorico.ativo_id,
                TickerHistorico.ticker,
                TickerHistorico.inicio,
                TickerHistorico.fim,
                Ativo.isin,
            )
            .join(Ativo, TickerHistorico.ativo_id == Ativo.id)
            .order_by(TickerHistorico.ticker, TickerHistorico.inicio)
        ).all()
    return [TickerCadastro(a, t, i, f, isin or "") for a, t, i, f, isin in linhas]
