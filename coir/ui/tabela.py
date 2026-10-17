"""Tabelas densas (AG Grid) com a formatação dos relatórios.

O texto de cada célula sai de `reports.modelo.formatar` (pt-BR, as mesmas casas do xlsx e do pdf);
o valor bruto vai junto, escondido, para ordenar números e datas e pintar negativos de vermelho.
Qualquer `Secao` de relatório vira tabela sem código próprio.
"""

from collections.abc import Callable, Sequence
from datetime import date
from decimal import Decimal
from typing import Any

from nicegui import ui

from coir.reports.modelo import NUMERICAS, Coluna, Secao, TipoColuna, Valor, formatar

Linha = Sequence[Valor]
POUCAS_LINHAS = 20


def _bruto(valor: Valor) -> str | None:
    if valor is None or isinstance(valor, str):
        return None
    if isinstance(valor, date):
        return valor.isoformat()
    return str(Decimal(valor))


def _coluna(i: int, c: Coluna, largura: int | None) -> dict[str, Any]:
    definicao: dict[str, Any] = {
        "headerName": c.nome,
        "field": f"c{i}",
        "tooltipField": f"c{i}",
        "sortable": True,
        "resizable": True,
    }
    if largura:
        definicao["width"] = largura
    if c.tipo in NUMERICAS:
        definicao |= {
            "type": "rightAligned",
            ":comparator": f"(a, b, x, y) => (+x.data.n{i} || 0) - (+y.data.n{i} || 0)",
            "cellClassRules": {"coir-negativo": f"Number(data.n{i}) < 0"},
        }
    elif c.tipo is TipoColuna.DATA:
        definicao[":comparator"] = (
            f"(a, b, x, y) => String(x.data.n{i} || '').localeCompare(String(y.data.n{i} || ''))"
        )
    return definicao


def _linha(colunas: Sequence[Coluna], linha: Linha, chave: object = None) -> dict[str, Any]:
    dados: dict[str, Any] = {"_chave": chave}
    for i, (c, v) in enumerate(zip(colunas, linha, strict=True)):
        dados[f"c{i}"] = formatar(v, c.tipo)
        dados[f"n{i}"] = _bruto(v)
    return dados


def linhas(
    colunas: Sequence[Coluna], valores: Sequence[Linha], chaves: Sequence[object] | None = None
) -> list[dict[str, Any]]:
    """As linhas no formato da grade (`chaves`: o que volta ao clicar numa linha)."""
    return [_linha(colunas, v, chaves[k] if chaves else None) for k, v in enumerate(valores)]


def grade(
    colunas: Sequence[Coluna],
    valores: Sequence[Linha],
    *,
    chaves: Sequence[object] | None = None,
    total: Linha | None = None,
    larguras: Sequence[int | None] | None = None,
    altura: str = "h-[520px]",
    ao_clicar: Callable[[object], object] | None = None,
) -> ui.aggrid:
    """Grade com ordenação, colunas redimensionáveis, total fixo no rodapé e clique na linha."""
    definicoes = [_coluna(i, c, larguras[i] if larguras else None) for i, c in enumerate(colunas)]
    opcoes: dict[str, Any] = {
        "columnDefs": definicoes,
        "rowData": linhas(colunas, valores, chaves),
        "rowSelection": {"mode": "singleRow", "checkboxes": False, "enableClickSelection": True},
        "rowHeight": 30,
        "headerHeight": 32,
        "tooltipShowDelay": 600,
        "suppressCellFocus": True,
        "pinnedBottomRowData": [_linha(colunas, total)] if total else [],
        "overlayNoRowsTemplate": "Nada para mostrar com estes filtros.",
    }
    if len(valores) <= POUCAS_LINHAS:  # a grade cresce com as linhas, sem área vazia
        opcoes["domLayout"] = "autoHeight"
        altura = ""
    tabela = ui.aggrid(opcoes, theme="balham", auto_size_columns=not larguras)
    tabela.classes(f"coir-grade w-full {altura}")
    if ao_clicar is not None:
        clicar = ao_clicar
        tabela.on("rowClicked", lambda e: clicar(e.args["data"]["_chave"]), args=["data"])
    return tabela


def atualizar(
    tabela: ui.aggrid,
    colunas: Sequence[Coluna],
    valores: Sequence[Linha],
    *,
    chaves: Sequence[object] | None = None,
    total: Linha | None = None,
) -> None:
    tabela.options["rowData"] = linhas(colunas, valores, chaves)
    tabela.options["pinnedBottomRowData"] = [_linha(colunas, total)] if total else []
    tabela.update()


def secao(s: Secao, altura: str = "h-[520px]") -> ui.aggrid:
    """Uma seção de relatório (reports.modelo) na tela."""
    if s.titulo:
        ui.label(s.titulo).classes("coir-titulo text-base")
    tabela = grade(s.colunas, s.linhas, total=s.total, altura=altura)
    if s.nota:
        ui.label(s.nota).classes("coir-subtitulo text-xs")
    return tabela
