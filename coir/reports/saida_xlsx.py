"""Relatórios -> um .xlsx (XlsxWriter; ADR 0003): uma aba por relatório, só valores, sem fórmulas,
compatível com o Excel 2016 (e LibreOffice), sem precisar do Excel para gerar.

Cada aba: título, titular e CPF, período e data de geração; depois cada seção (título, cabeçalho,
linhas e total). Números com formato pela coluna; negativos em vermelho.

Decimal: o xlsx guarda número como double; o valor é arredondado à precisão da coluna e só
então convertido (único ponto com float, na gravação).
"""

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import xlsxwriter
from xlsxwriter.format import Format
from xlsxwriter.worksheet import Worksheet

from coir.reports.modelo import (
    NUMERICAS,
    Relatorio,
    Secao,
    TipoColuna,
    Valor,
    arredondar,
    cpf_formatado,
    formatar,
)

FORMATOS = {
    TipoColuna.INTEIRO: "#,##0;[Red]-#,##0",
    TipoColuna.QUANTIDADE: "#,##0.########;[Red]-#,##0.########",
    TipoColuna.DINHEIRO: "#,##0.00;[Red]-#,##0.00",
    TipoColuna.PM: "#,##0.00000000;[Red]-#,##0.00000000",
    TipoColuna.PERCENTUAL: "0.00%;[Red]-0.00%",
    TipoColuna.DATA: "dd/mm/yyyy",
}
LARGURA_MAX = 60


def _nome_aba(titulo: str, usados: set[str]) -> str:
    base = "".join(c for c in titulo if c not in "[]:*?/\\")[:31]
    nome, n = base, 2
    while nome.lower() in usados:
        sufixo = f" ({n})"
        nome, n = base[: 31 - len(sufixo)] + sufixo, n + 1
    usados.add(nome.lower())
    return nome


class _Formatos:
    def __init__(self, livro: xlsxwriter.Workbook) -> None:
        self.titulo = livro.add_format({"bold": True, "font_size": 14})
        self.info = livro.add_format({"italic": True, "font_color": "#555555"})
        self.secao = livro.add_format({"bold": True, "font_size": 11})
        self.cabecalho = livro.add_format({"bold": True, "bg_color": "#DDE4EE", "border": 1,
                                           "text_wrap": True, "valign": "top"})  # fmt: skip
        self.nota = livro.add_format({"italic": True, "font_size": 9, "font_color": "#555555"})
        self._livro = livro
        self._cache: dict[tuple[TipoColuna, bool], Format] = {}

    def celula(self, tipo: TipoColuna, total: bool) -> Format:
        chave = (tipo, total)
        if chave not in self._cache:
            props: dict[str, object] = {"bold": total}
            if tipo in FORMATOS:
                props["num_format"] = FORMATOS[tipo]
            if total:
                props["top"] = 1
            self._cache[chave] = self._livro.add_format(props)
        return self._cache[chave]


def _gravar(aba: Worksheet, lin: int, col: int, valor: Valor, tipo: TipoColuna, fmt: Format
            ) -> None:  # fmt: skip
    if valor is None or valor == "":
        aba.write_blank(lin, col, None, fmt)
    elif isinstance(valor, (Decimal, int)) and tipo in NUMERICAS:
        aba.write_number(lin, col, float(arredondar(Decimal(valor), tipo)), fmt)
    elif isinstance(valor, datetime):
        aba.write_string(lin, col, formatar(valor, tipo), fmt)
    elif isinstance(valor, date):
        aba.write_datetime(lin, col, datetime(valor.year, valor.month, valor.day), fmt)
    else:
        aba.write_string(lin, col, formatar(valor, tipo), fmt)


def _secao(aba: Worksheet, lin: int, s: Secao, f: _Formatos, larguras: dict[int, int]) -> int:
    aba.write_string(lin, 0, s.titulo, f.secao)
    lin += 1
    for c, coluna in enumerate(s.colunas):
        aba.write_string(lin, c, coluna.nome, f.cabecalho)
        larguras[c] = max(larguras.get(c, 8), min(len(coluna.nome) + 2, 24))
    lin += 1
    for linha, total in [*((x, False) for x in s.linhas), *(((s.total, True),) if s.total else ())]:
        for c, (valor, coluna) in enumerate(zip(linha, s.colunas, strict=True)):
            _gravar(aba, lin, c, valor, coluna.tipo, f.celula(coluna.tipo, total))
            larguras[c] = max(larguras[c], min(len(formatar(valor, coluna.tipo)) + 2, LARGURA_MAX))
        lin += 1
    if s.nota:
        aba.write_string(lin, 0, s.nota, f.nota)
        lin += 1
    return lin + 1


def escrever_xlsx(relatorios: list[Relatorio], caminho: Path, agora: datetime) -> Path:
    livro = xlsxwriter.Workbook(str(caminho), {"strings_to_numbers": False,
                                               "strings_to_formulas": False})  # fmt: skip
    livro.set_properties({"title": "COIR - relatórios fiscais", "author": "COIR",
                          "created": agora})  # fmt: skip
    f = _Formatos(livro)
    usados: set[str] = set()
    for rel in relatorios:
        aba = livro.add_worksheet(_nome_aba(rel.aba or rel.titulo, usados))
        cab = rel.cabecalho
        aba.write_string(0, 0, rel.titulo, f.titulo)
        aba.write_string(1, 0, f"{cab.titular} - CPF {cpf_formatado(cab.cpf)}", f.info)
        aba.write_string(2, 0, f"{cab.periodo} - gerado em {cab.gerado_em:%d/%m/%Y %H:%M}", f.info)
        lin = 4
        larguras: dict[int, int] = {}
        if len(rel.secoes) == 1:  # tabela única: cabeçalho congelado e filtro
            s = rel.secoes[0]
            aba.freeze_panes(lin + 2, 0)
            ultima = lin + 1 + len(s.linhas)
            aba.autofilter(lin + 1, 0, max(ultima, lin + 1), len(s.colunas) - 1)
        for s in rel.secoes:
            lin = _secao(aba, lin, s, f, larguras)
        for nota in rel.notas:
            aba.write_string(lin, 0, nota, f.nota)
            lin += 1
        for c, largura in larguras.items():
            aba.set_column(c, c, largura)
    livro.close()
    return caminho
