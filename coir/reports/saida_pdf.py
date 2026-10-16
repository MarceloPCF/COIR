"""Relatório -> PDF (reportlab; ADR 0003): A4 paisagem, cabeçalho com título, titular, CPF,
período e data de geração em cada página; rodapé "Página X de Y"; tabelas longas quebram de
página repetindo o cabeçalho. Fontes padrão Helvetica (WinAnsi cobre os acentos do português).
"""

from decimal import Decimal
from pathlib import Path
from typing import Any

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import LongTable, Paragraph, SimpleDocTemplate, Spacer, TableStyle
from reportlab.platypus.flowables import Flowable

from coir.reports.modelo import NUMERICAS, Relatorio, Secao, cpf_formatado, formatar

PAGINA = landscape(A4)
MARGEM = 12 * mm
FONTE, NEGRITO, CORPO = "Helvetica", "Helvetica-Bold", 7.5
AVISO = ("Relatório gerado pelo COIR a partir das notas importadas e dos eventos lançados. "
         "Confira antes de declarar; não substitui a orientação de um contador.")  # fmt: skip
ESTILO_SECAO = ParagraphStyle(
    "secao", fontName=NEGRITO, fontSize=10, spaceBefore=6, spaceAfter=3, keepWithNext=1
)
CINZA = colors.HexColor("#555555")
ESTILO_NOTA = ParagraphStyle("nota", fontName=FONTE, fontSize=7, textColor=CINZA)


def _escapar(texto: str) -> str:
    return texto.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


class _Paginas(Canvas):
    """Canvas que conhece o total de páginas (cabeçalho e rodapé desenhados no fim)."""

    def __init__(self, *args: Any, rel: Relatorio, **kwargs: Any) -> None:
        super().__init__(*args, **kwargs)
        self._rel = rel
        self._paginas: list[dict[str, Any]] = []

    def showPage(self) -> None:  # nome da API do reportlab
        self._paginas.append(dict(self.__dict__))
        self._startPage()  # pyright: ignore[reportAttributeAccessIssue]  (receita do reportlab)

    def save(self) -> None:
        total = len(self._paginas)
        for estado in self._paginas:
            self.__dict__.update(estado)
            self._moldura(total)
            super().showPage()
        super().save()

    def _moldura(self, total: int) -> None:
        cab = self._rel.cabecalho
        largura, altura = PAGINA
        self.setFont(NEGRITO, 12)
        self.drawString(MARGEM, altura - MARGEM, self._rel.titulo)
        self.setFont(FONTE, 8)
        self.drawString(
            MARGEM,
            altura - MARGEM - 12,
            f"{cab.titular}  -  CPF {cpf_formatado(cab.cpf)}  -  {cab.periodo}",
        )
        self.drawRightString(largura - MARGEM, altura - MARGEM,
                             f"Gerado em {cab.gerado_em:%d/%m/%Y %H:%M}")  # fmt: skip
        self.setStrokeColor(colors.HexColor("#999999"))
        self.line(MARGEM, altura - MARGEM - 17, largura - MARGEM, altura - MARGEM - 17)
        self.setFont(FONTE, 7)
        self.drawString(MARGEM, MARGEM - 6, "COIR - Controle de Operações e Imposto de Renda")
        self.drawRightString(largura - MARGEM, MARGEM - 6,
                             f"Página {self.getPageNumber()} de {total}")  # fmt: skip


def _larguras(s: Secao, celulas: list[list[str]], disponivel: float) -> list[float]:
    """Largura pelo conteúdo; o título da coluna quebra linha (vale a maior palavra dele)."""
    medidas = []
    for c, coluna in enumerate(s.colunas):
        palavra = max(stringWidth(w, NEGRITO, CORPO) for w in coluna.nome.split())
        maior = max([palavra] + [stringWidth(linha[c], FONTE, CORPO) for linha in celulas])
        medidas.append(min(maior + 10, 180.0))  # 6 de padding + folga
    soma = sum(medidas)
    return medidas if soma <= disponivel else [m * disponivel / soma for m in medidas]


def _quebrar(texto: str, largura: float) -> str:
    """Título da coluna em linhas que cabem na largura, sem partir palavra."""
    linhas: list[str] = []
    for palavra in texto.split():
        tentativa = f"{linhas[-1]} {palavra}" if linhas else palavra
        if linhas and stringWidth(tentativa, NEGRITO, CORPO) <= largura:
            linhas[-1] = tentativa
        else:
            linhas.append(palavra)
    return "\n".join(linhas)


def _tabela(s: Secao, disponivel: float) -> LongTable:
    linhas = [*s.linhas, *((s.total,) if s.total else ())]
    celulas = [[formatar(v, c.tipo) for v, c in zip(linha, s.colunas, strict=True)]
               for linha in linhas]  # fmt: skip
    larguras = _larguras(s, celulas, disponivel)
    titulos = [_quebrar(c.nome, w - 6) for c, w in zip(s.colunas, larguras, strict=True)]
    dados: list[list[Any]] = [titulos, *celulas]
    estilo: list[tuple[Any, ...]] = [
        ("FONT", (0, 0), (-1, -1), FONTE, CORPO),
        ("FONT", (0, 0), (-1, 0), NEGRITO, CORPO),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#DDE4EE")),
        ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.HexColor("#999999")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F7FA")]),
    ]
    for c, coluna in enumerate(s.colunas):
        if coluna.tipo in NUMERICAS:
            estilo.append(("ALIGN", (c, 0), (c, -1), "RIGHT"))
    for i, linha in enumerate(linhas, start=1):
        for c, valor in enumerate(linha):
            if isinstance(valor, (Decimal, int)) and valor < 0:
                estilo.append(("TEXTCOLOR", (c, i), (c, i), colors.HexColor("#B00020")))
    if s.total:
        estilo += [("FONT", (0, -1), (-1, -1), NEGRITO, CORPO),
                   ("LINEABOVE", (0, -1), (-1, -1), 0.5, colors.black)]  # fmt: skip
    tabela = LongTable(dados, colWidths=larguras, repeatRows=1, hAlign="LEFT")
    tabela.setStyle(TableStyle(estilo))
    return tabela


def escrever_pdf(rel: Relatorio, caminho: Path) -> Path:
    """PDF invariante (sem data de criação embutida): a data do relatório está no cabeçalho."""
    doc = SimpleDocTemplate(
        str(caminho), pagesize=PAGINA, leftMargin=MARGEM, rightMargin=MARGEM,
        topMargin=MARGEM + 22, bottomMargin=MARGEM + 4, title=rel.titulo, author="COIR",
        creator="COIR", invariant=1,
    )  # fmt: skip
    disponivel = PAGINA[0] - 2 * MARGEM
    historia: list[Flowable] = []
    for s in rel.secoes:
        historia.append(Paragraph(_escapar(s.titulo), ESTILO_SECAO))
        if s.linhas or s.total:
            historia.append(_tabela(s, disponivel))
        else:
            historia.append(Paragraph("Sem movimento no período.", ESTILO_NOTA))
        if s.nota:
            historia += [Spacer(1, 2), Paragraph(_escapar(s.nota), ESTILO_NOTA)]
    historia.append(Spacer(1, 6))
    for nota in (*rel.notas, AVISO):
        historia.append(Paragraph(_escapar(nota), ESTILO_NOTA))
    doc.build(historia, canvasmaker=lambda *a, **k: _Paginas(*a, rel=rel, **k))  # type: ignore[arg-type]
    return caminho
