"""Tela 14 (declaração anual): marcação dos itens lançados no programa IRPF e a conferência com
a declaração entregue (PDF em ``<dados>/_dirpf/``)."""

from collections import defaultdict
from collections.abc import Callable, Iterable, Mapping
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from itertools import combinations
from pathlib import Path

from sqlalchemy import select

from coir.db.modelos import IrpfLancado
from coir.db.sessao import conectar
from coir.domain.irpf import FichasIrpf, MapeamentoIrpf
from coir.domain.proventos import Provento, TipoProvento
from coir.erros import CoirError
from coir.gabarito.comparar import Tolerancias
from coir.gabarito.dirpf import Dirpf, ler_dirpf
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


# ------------------------------------------------- crédito do JCP proposto pela DIRPF entregue
MAX_PARCELAS = 8  # pagamentos de um mesmo crédito (parcelas) procurados por combinação
TOLERANCIA = Decimal("0.01")


@dataclass(frozen=True)
class PropostaCredito:
    """JCP pago num ano cujo crédito, pela DIRPF entregue, foi no ano anterior (Bens e Direitos
    99.07, JCP a receber em 31/12). A data proposta é 31/12 do ano do crédito: a DIRPF só diz o
    ano, que é o que decide a declaração."""

    chave: str  # da linha em movimentacao_b3
    codigo: str
    pagamento: date
    liquido: Decimal
    cnpj: str
    credito: date
    origem: str  # PDF que deu a proposta


def _a_receber(
    declaracoes: Iterable[tuple[str, Dirpf]],
) -> dict[tuple[str, int], tuple[Decimal, str]]:
    """(CNPJ, ano do crédito) -> (valor a receber em 31/12, PDF), das colunas das DIRPF lidas."""
    saida: dict[tuple[str, int], tuple[Decimal, str]] = {}
    for nome, d in declaracoes:
        for b in d.bens:
            if (b.grupo, b.codigo) != ("99", "07") or not b.cnpj:
                continue
            for ano, valor in ((d.ano - 1, b.anterior), (d.ano, b.atual)):
                if valor:
                    atual = saida.get((b.cnpj, ano), (Decimal(0), nome))
                    saida[(b.cnpj, ano)] = (atual[0] + valor, nome) if atual[1] == nome else atual
    return saida


def _fecha(pagos: Iterable[Provento], valor: Decimal) -> bool:
    return abs(sum((p.liquido for p in pagos), Decimal(0)) - valor) <= TOLERANCIA


def propor_creditos(
    proventos: Iterable[Provento], cnpj_de: Mapping[str, str],
    declaracoes: Iterable[tuple[str, Dirpf]],
) -> list[PropostaCredito]:  # fmt: skip
    """Para cada JCP a receber em 31/12 declarado, os JCP pagos no ano seguinte pela mesma empresa
    cuja soma fecha o valor (uma combinação só; mais de uma ou nenhuma: sem proposta). JCP com a
    data do crédito informada pelo usuário fica de fora."""
    por_empresa: dict[tuple[str, int], list[Provento]] = defaultdict(list)
    for p in proventos:
        cnpj = cnpj_de.get(p.codigo)
        if p.tipo is TipoProvento.JCP and cnpj and "(informada)" not in p.observacao:
            por_empresa[(cnpj, p.data.year)].append(p)
    saida: list[PropostaCredito] = []
    for (cnpj, ano), (valor, origem) in sorted(_a_receber(declaracoes).items()):
        pagos = sorted(por_empresa.get((cnpj, ano + 1), []), key=lambda p: (p.data, p.chave))
        fecham = [c for n in range(1, min(len(pagos), MAX_PARCELAS) + 1)
                  for c in combinations(pagos, n) if _fecha(c, valor)]  # fmt: skip
        if len(fecham) != 1:
            continue
        saida += [PropostaCredito(chave=p.chave, codigo=p.codigo, pagamento=p.data,
                                  liquido=p.liquido, cnpj=cnpj, credito=date(ano, 12, 31),
                                  origem=origem)
                  for p in fecham[0] if p.ano_declaracao != ano]  # fmt: skip
    return saida


def declaracoes_lidas(dados: Path) -> tuple[list[tuple[str, Dirpf]], list[str]]:
    """As DIRPF de <dados>/_dirpf que o COIR consegue ler (as outras viram aviso)."""
    lidas: list[tuple[str, Dirpf]] = []
    avisos: list[str] = []
    for pdf in pdfs_dirpf(dados):
        try:
            lidas.append((pdf.name, ler_dirpf(pdf)))
        except CoirError as erro:
            avisos.append(f"{pdf.name}: não lida ({erro})")
    return lidas, avisos


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
