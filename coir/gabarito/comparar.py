"""Comparação candidato x gabarito com tolerâncias.

Três níveis:
- ``operacoes``: somas por (corretora, conta, nota, data, papel, C/V) — imune a diferenças de
  consolidação (o gabarito às vezes junta, às vezes separa, linhas da mesma nota);
- ``posicoes``: PM e quantidade após cada operação, na ordem, por (categoria, conta, papel);
- ``apuracao``: valores mensais por categoria.
"""

import tomllib
from collections import defaultdict
from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path

from coir.erros import GabaritoError
from coir.gabarito.modelo import ApuracaoMensalGabarito, Gabarito, OperacaoGabarito, PosicaoGabarito
from coir.recursos import pasta_data

ChaveOperacao = tuple[str, str, str, date, str, str]


@dataclass(frozen=True)
class Tolerancias:
    operacao: Decimal
    pm: Decimal
    mes: Decimal

    @classmethod
    def carregar(cls, caminho: Path | None = None) -> "Tolerancias":
        arquivo = caminho or pasta_data() / "tolerancias.toml"
        try:
            dados = tomllib.loads(arquivo.read_text(encoding="utf-8"))
            return cls(**{k: Decimal(str(dados[k])) for k in ("operacao", "pm", "mes")})
        except (KeyError, ArithmeticError, tomllib.TOMLDecodeError) as erro:
            raise GabaritoError(f"Arquivo de tolerâncias inválido: {arquivo}") from erro


@dataclass(frozen=True)
class Divergencia:
    nivel: str
    chave: str
    campo: str
    esperado: Decimal | None
    obtido: Decimal | None

    @property
    def diferenca(self) -> Decimal | None:
        if self.esperado is None or self.obtido is None:
            return None
        return self.obtido - self.esperado


@dataclass
class Relatorio:
    comparados: dict[str, int] = field(default_factory=dict)
    divergencias: list[Divergencia] = field(default_factory=list)
    faltando: list[str] = field(default_factory=list)  # no gabarito, não no candidato
    sobrando: list[str] = field(default_factory=list)  # no candidato, não no gabarito

    @property
    def ok(self) -> bool:
        return not (self.divergencias or self.faltando or self.sobrando)


@dataclass
class _Somas:
    quantidade: Decimal = Decimal(0)
    total: Decimal = Decimal(0)
    custos: Decimal = Decimal(0)
    irrf: Decimal = Decimal(0)


def _chave_operacao(op: OperacaoGabarito) -> ChaveOperacao:
    return (op.corretora, op.conta, op.nota, op.data, op.papel, op.cv)


def somar_operacoes(operacoes: Iterable[OperacaoGabarito]) -> dict[ChaveOperacao, _Somas]:
    somas: dict[ChaveOperacao, _Somas] = defaultdict(_Somas)
    for op in operacoes:
        s = somas[_chave_operacao(op)]
        s.quantidade += op.quantidade
        s.total += op.total
        s.custos += op.custos
        s.irrf += op.irrf
    return dict(somas)


def _fora(esperado: Decimal, obtido: Decimal, tolerancia: Decimal) -> bool:
    return abs(obtido - esperado) > tolerancia


def comparar_operacoes(
    esperado: Iterable[OperacaoGabarito],
    obtido: Iterable[OperacaoGabarito],
    tol: Tolerancias,
    rel: Relatorio,
) -> None:
    a, b = somar_operacoes(esperado), somar_operacoes(obtido)
    rel.comparados["operacoes"] = len(a.keys() & b.keys())
    rel.faltando += [f"operacao {k}" for k in sorted(a.keys() - b.keys(), key=str)]
    rel.sobrando += [f"operacao {k}" for k in sorted(b.keys() - a.keys(), key=str)]
    for chave in sorted(a.keys() & b.keys(), key=str):
        e, o = a[chave], b[chave]
        if e.quantidade != o.quantidade:
            rel.divergencias.append(
                Divergencia("operacao", str(chave), "quantidade", e.quantidade, o.quantidade)
            )
        for campo in ("total", "custos", "irrf"):
            ve, vo = getattr(e, campo), getattr(o, campo)
            if _fora(ve, vo, tol.operacao):
                rel.divergencias.append(Divergencia("operacao", str(chave), campo, ve, vo))


def _sequencias(
    posicoes: Iterable[PosicaoGabarito],
) -> dict[tuple[str, str, str], list[PosicaoGabarito]]:
    seq: dict[tuple[str, str, str], list[PosicaoGabarito]] = defaultdict(list)
    for p in sorted(posicoes, key=lambda p: p.ordem):
        seq[(p.categoria.value, p.conta, p.papel)].append(p)
    return dict(seq)


def comparar_posicoes(
    esperado: Iterable[PosicaoGabarito],
    obtido: Iterable[PosicaoGabarito],
    tol: Tolerancias,
    rel: Relatorio,
) -> None:
    a, b = _sequencias(esperado), _sequencias(obtido)
    total = 0
    for chave in sorted(a.keys() | b.keys()):
        lista_a, lista_b = a.get(chave, []), b.get(chave, [])
        for i in range(max(len(lista_a), len(lista_b))):
            rotulo = f"{chave} #{i + 1}"
            if i >= len(lista_b):
                rel.faltando.append(f"posicao {rotulo}")
                continue
            if i >= len(lista_a):
                rel.sobrando.append(f"posicao {rotulo}")
                continue
            pa, pb = lista_a[i], lista_b[i]
            total += 1
            if pa.quantidade_apos != pb.quantidade_apos:
                rel.divergencias.append(
                    Divergencia(
                        "posicao", rotulo, "quantidade_apos", pa.quantidade_apos, pb.quantidade_apos
                    )
                )
            if (pa.pm_apos is None) != (pb.pm_apos is None) or (
                pa.pm_apos is not None
                and pb.pm_apos is not None
                and _fora(pa.pm_apos, pb.pm_apos, tol.pm)
            ):
                rel.divergencias.append(
                    Divergencia("posicao", rotulo, "pm_apos", pa.pm_apos, pb.pm_apos)
                )
    rel.comparados["posicoes"] = total


CAMPOS_MES = ("resultado", "irrf", "a_pagar", "acumulado", "base_calculo")


def comparar_apuracao(
    esperado: Iterable[ApuracaoMensalGabarito],
    obtido: Iterable[ApuracaoMensalGabarito],
    tol: Tolerancias,
    rel: Relatorio,
) -> None:
    a = {(x.categoria.value, x.ano, x.mes): x for x in esperado}
    b = {(x.categoria.value, x.ano, x.mes): x for x in obtido}
    rel.comparados["apuracao"] = len(a.keys() & b.keys())
    rel.faltando += [f"apuracao {k}" for k in sorted(a.keys() - b.keys())]
    rel.sobrando += [f"apuracao {k}" for k in sorted(b.keys() - a.keys())]
    for chave in sorted(a.keys() & b.keys()):
        for campo in CAMPOS_MES:
            ve, vo = getattr(a[chave], campo), getattr(b[chave], campo)
            if _fora(ve, vo, tol.mes):
                rel.divergencias.append(Divergencia("apuracao", str(chave), campo, ve, vo))


NIVEIS = ("operacoes", "posicoes", "apuracao")


def comparar(
    esperado: Gabarito, obtido: Gabarito, tol: Tolerancias, niveis: Iterable[str] = NIVEIS
) -> Relatorio:
    rel = Relatorio()
    for nivel in niveis:
        if nivel == "operacoes":
            comparar_operacoes(esperado.operacoes, obtido.operacoes, tol, rel)
        elif nivel == "posicoes":
            comparar_posicoes(esperado.posicoes, obtido.posicoes, tol, rel)
        elif nivel == "apuracao":
            comparar_apuracao(esperado.apuracao, obtido.apuracao, tol, rel)
        else:
            raise GabaritoError(f"Nível de comparação desconhecido: {nivel}")
    return rel
