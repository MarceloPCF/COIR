"""Notas gravadas para a interface: a lista com filtros, o detalhe com a conciliação e a página
do PDF em imagem (para conferir a nota ao lado do original)."""

from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from io import BytesIO
from pathlib import Path

import pypdfium2 as pdfium  # pyright: ignore[reportMissingTypeStubs]
from sqlalchemy import select
from sqlalchemy.orm import selectinload

from coir.db.modelos import ArquivoImportado, Conta, NotaCorretagem
from coir.db.repositorio import operacoes_gravadas
from coir.db.sessao import conectar
from coir.domain.nota import NotaExtraida, OperacaoExtraida
from coir.erros import BancoError, ConciliacaoError, ExtracaoError
from coir.parsing.conciliar import conciliar

PASTA_NOTAS = "_notas-corretagens"


@dataclass(frozen=True)
class ResumoNota:
    id: int
    data: date
    corretora: str
    conta: str
    numero: str
    segmento: str
    valor_operacoes: Decimal
    custos: Decimal
    irrf: Decimal
    liquido: Decimal
    arquivo: str
    paginas: tuple[int, ...]


@dataclass(frozen=True)
class Conferencia:
    rotulo: str
    impresso: Decimal
    calculado: Decimal | None = None  # None: só o valor impresso

    @property
    def confere(self) -> bool:
        return self.calculado is None or self.calculado == self.impresso


@dataclass(frozen=True)
class DetalheNota:
    resumo: ResumoNota
    nota: NotaExtraida
    operacoes: tuple[OperacaoExtraida, ...]
    conferencias: tuple[Conferencia, ...]
    conciliada: bool
    motivo: str  # por que não concilia (vazio se concilia)
    pdf: Path | None  # o arquivo na pasta das notas, se estiver lá


def _resumo(n: NotaCorretagem) -> ResumoNota:
    return ResumoNota(
        id=n.id, data=n.data_pregao, corretora=n.conta.corretora.nome, conta=n.conta.numero,
        numero=n.numero, segmento=n.segmento, valor_operacoes=n.valor_operacoes,
        custos=n.custos_total, irrf=n.irrf + n.irrf_day_trade, liquido=n.liquido,
        arquivo=n.arquivo.nome, paginas=tuple(int(p) for p in n.paginas.split(",") if p),
    )  # fmt: skip


def listar(banco: Path, titular: int | None, ano: int | None) -> list[ResumoNota]:
    """Notas do titular no ano, da mais recente à mais antiga."""
    consulta = (
        select(NotaCorretagem)
        .join(Conta)
        .options(
            selectinload(NotaCorretagem.conta).selectinload(Conta.corretora),
            selectinload(NotaCorretagem.arquivo),
        )
        .order_by(NotaCorretagem.data_pregao.desc(), NotaCorretagem.id.desc())
    )
    if titular is not None:
        consulta = consulta.where(Conta.titular_id == titular)
    if ano is not None:
        no_ano = NotaCorretagem.data_pregao.between(date(ano, 1, 1), date(ano, 12, 31))
        consulta = consulta.where(no_ano)
    with conectar(banco) as fabrica, fabrica() as s:
        return [_resumo(n) for n in s.scalars(consulta)]


def _conferencias(nota: NotaExtraida) -> tuple[Conferencia, ...]:
    if nota.resumo_bovespa is not None:
        r = nota.resumo_bovespa
        negocios = sum((n.valor for n in nota.negocios_bovespa), Decimal(0))
        return (
            Conferencia("Valor das operações", r.valor_operacoes, negocios),
            Conferencia("Compras", r.compras),
            Conferencia("Vendas", r.vendas),
            Conferencia("Custos (sem IRRF)", nota.custos_total),
            Conferencia("IRRF", r.irrf + r.irrf_day_trade),
            Conferencia("Líquido", r.liquido.assinado),
        )
    r2 = nota.resumo_bmf
    if r2 is None:
        raise ExtracaoError(f"nota {nota.numero} sem resumo")
    negocios = sum((n.valor if n.dc == "C" else -n.valor for n in nota.negocios_bmf), Decimal(0))
    return (
        Conferencia("Valor dos negócios", r2.valor_negocios.assinado, negocios),
        Conferencia("Custos (com impostos)", nota.custos_total),
        Conferencia("IRRF", r2.irrf + r2.irrf_day_trade),
        Conferencia("Líquido", r2.liquido.assinado),
    )


def detalhe(banco: Path, pasta_notas: Path, id_nota: int) -> DetalheNota:
    """A nota como impressa, as operações normalizadas e a conciliação refeita (regra 4)."""
    with conectar(banco) as fabrica, fabrica() as s:
        registro = s.scalars(
            select(NotaCorretagem)
            .where(NotaCorretagem.id == id_nota)
            .options(
                selectinload(NotaCorretagem.conta).selectinload(Conta.corretora),
                selectinload(NotaCorretagem.arquivo),
                selectinload(NotaCorretagem.operacoes),
            )
        ).one_or_none()
        if registro is None:
            raise BancoError(f"nota {id_nota} não existe no banco")
        resumo, ops = _resumo(registro), tuple(operacoes_gravadas(registro))
        nota = NotaExtraida.model_validate_json(registro.json_nota)
    try:
        conciliar(nota)
        conciliada, motivo = True, ""
    except ConciliacaoError as erro:
        conciliada, motivo = False, str(erro)
    pdf = pasta_notas / resumo.arquivo
    return DetalheNota(resumo, nota, ops, _conferencias(nota), conciliada, motivo,
                       pdf if pdf.is_file() else None)  # fmt: skip


def pagina_png(pdf: Path, pagina: int, largura: int = 1100) -> bytes:
    """Página (1..n) do PDF em PNG, com `largura` pixels."""
    try:
        documento = pdfium.PdfDocument(pdf)
    except (OSError, pdfium.PdfiumError) as erro:
        raise ExtracaoError(f"não foi possível abrir {pdf.name}: {erro}") from erro
    try:
        if not 1 <= pagina <= len(documento):
            raise ExtracaoError(f"{pdf.name}: página {pagina} não existe ({len(documento)})")
        folha = documento[pagina - 1]
        escala = largura / folha.get_width()
        imagem = folha.render(scale=escala).to_pil()
        saida = BytesIO()
        imagem.save(saida, format="PNG", optimize=True)
        return saida.getvalue()
    finally:
        documento.close()


def arquivos_incompletos(banco: Path) -> list[str]:
    """PDFs com nota rejeitada na última importação (relidos na próxima)."""
    with conectar(banco) as fabrica, fabrica() as s:
        consulta = select(ArquivoImportado.nome).where(ArquivoImportado.completo.is_(False))
        return sorted(s.scalars(consulta))
