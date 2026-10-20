"""Tela 14 (declaração anual): marcação dos itens lançados no programa IRPF e a conferência com
a declaração entregue (PDF em ``<dados>/_dirpf/``)."""

from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from sqlalchemy import select

from coir.db.modelos import IrpfLancado
from coir.db.sessao import conectar
from coir.domain.irpf import FichasIrpf, MapeamentoIrpf
from coir.gabarito.comparar import Tolerancias
from coir.gabarito.dirpf import ler_dirpf
from coir.gabarito.dirpf_comparar import LinhaComparacao, carregar_explicacoes, comparar, explicar

PASTA_DIRPF = "_dirpf"


def lancados(banco: Path, titular: int, ano: int) -> set[str]:
    with conectar(banco) as fabrica, fabrica() as s:
        consulta = select(IrpfLancado.chave).where(IrpfLancado.titular_id == titular,
                                                   IrpfLancado.ano == ano)  # fmt: skip
        return set(s.scalars(consulta))


def marcar(banco: Path, titular: int, ano: int, chave: str, lancado: bool,
           agora: datetime) -> None:  # fmt: skip
    """Marca (ou desmarca) um item como lançado no programa IRPF."""
    with conectar(banco) as fabrica, fabrica() as s, s.begin():
        atual = s.scalar(select(IrpfLancado).where(
            IrpfLancado.titular_id == titular, IrpfLancado.ano == ano,
            IrpfLancado.chave == chave))  # fmt: skip
        if lancado and atual is None:
            s.add(IrpfLancado(titular_id=titular, ano=ano, chave=chave, marcado_em=agora))
        elif not lancado and atual is not None:
            s.delete(atual)


@dataclass(frozen=True)
class Conferencia:
    arquivo: str  # nome do PDF lido (sem caminho)
    linhas: tuple[LinhaComparacao, ...]

    @property
    def abertas(self) -> tuple[LinhaComparacao, ...]:
        return tuple(x for x in self.linhas if x.classe not in ("confere", "explicada"))


def pdfs_dirpf(dados: Path) -> list[Path]:
    return sorted((dados / PASTA_DIRPF).glob("*.pdf"))


def conferir(
    dados: Path, fichas: FichasIrpf, mapa: MapeamentoIrpf, codigo_atual: Callable[[str], str],
) -> Conferencia | None:  # fmt: skip
    """Compara as fichas com a DIRPF entregue do mesmo ano-calendário, se houver o PDF em
    ``<dados>/_dirpf``; as divergências já analisadas (``gabarito/dirpf-explicadas.csv``) saem
    como explicadas."""
    for pdf in pdfs_dirpf(dados):
        declarada = ler_dirpf(pdf)
        if declarada.ano == fichas.ano:
            tol = Tolerancias.carregar()
            explicacoes = carregar_explicacoes(dados / "gabarito" / "dirpf-explicadas.csv")
            linhas = explicar(comparar(fichas, declarada, mapa, tol.mes, tol.operacao,
                                       codigo_atual), explicacoes)  # fmt: skip
            return Conferencia(pdf.name, tuple(linhas))
    return None
