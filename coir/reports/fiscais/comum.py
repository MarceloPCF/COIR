"""Nomes, meses e formatação comuns aos relatórios fiscais (cabeçalho, totais, categorias)."""

from collections.abc import Iterable
from decimal import Decimal

from coir.domain.fiscal import (
    ZERO,
    Categoria,
    Grupo,
    MesApurado,
    Movimento,
    ResultadoCategoria,
    TipoAtivo,
)
from coir.domain.relatorio import DadosRelatorio
from coir.reports.modelo import (
    Cabecalho,
    Coluna,
    TipoColuna,
    Valor,
    arredondar,
)

T = TipoColuna


MESES = (
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro",
)


NOME_CATEGORIA = {
    Categoria.COMUM: "Operações comuns",
    Categoria.DAY_TRADE: "Day trade",
    Categoria.FII: "Fundos imobiliários e Fiagro",
}


NOME_GRUPO = {
    Grupo.A_VISTA: "À vista",
    Grupo.OPCOES: "Opções",
    Grupo.OPCOES_OUTROS: "Opções outros",
    Grupo.DOLAR: "Futuro dólar",
    Grupo.INDICE: "Futuro índices",
    Grupo.JUROS: "Futuro juros",
    Grupo.COMMODITY: "Futuro outros",
    Grupo.TERMO: "Termo",
}


CATEGORIA_CURTA = {Categoria.COMUM: "Comum", Categoria.DAY_TRADE: "Day trade",
                   Categoria.FII: "FII"}  # fmt: skip


GRUPOS = (Grupo.A_VISTA, Grupo.OPCOES, Grupo.OPCOES_OUTROS, Grupo.DOLAR, Grupo.INDICE,
          Grupo.JUROS, Grupo.COMMODITY, Grupo.TERMO)  # fmt: skip


def _dinheiro(nome: str) -> Coluna:
    return Coluna(nome, T.DINHEIRO)


def _cab(d: DadosRelatorio, periodo: str | None = None) -> Cabecalho:
    if periodo is None:
        periodo = f"Ano-calendário {d.ano}"
        if (d.ate.month, d.ate.day) != (12, 31):
            periodo += f" (parcial, até {d.ate:%d/%m/%Y})"
    return Cabecalho(d.nome, d.cpf, periodo, d.gerado_em)


def _meses(d: DadosRelatorio) -> list[MesApurado]:
    return sorted(
        (m for m in d.apuracao.meses if m.titular == d.titular and m.ano == d.ano),
        key=lambda m: m.mes,
    )


def _cat(m: MesApurado, categoria: Categoria) -> ResultadoCategoria:
    return next(c for c in m.categorias if c.categoria is categoria)


def _ativa(meses: list[MesApurado], categoria: Categoria) -> bool:
    """Categoria com algum valor no ano (resultado, imposto ou prejuízo de anos anteriores)."""
    return any(
        r.resultado
        or r.imposto
        or r.prejuizo_anterior
        or r.prejuizo_a_compensar
        or r.isento
        or r.vendas_limite
        or any(r.por_grupo.values())
        for r in (_cat(m, categoria) for m in meses)
    )


def _movimentos(d: DadosRelatorio) -> list[Movimento]:
    return [m for m in d.apuracao.movimentos if m.titular == d.titular and m.data.year == d.ano]


def _soma(valores: Iterable[Decimal]) -> Decimal:
    return sum(valores, ZERO)


def _total(
    linhas: list[tuple[Valor, ...]], colunas: tuple[Coluna, ...], somar: set[int]
) -> tuple[Valor, ...]:
    def coluna(i: int) -> Valor:
        if i == 0:
            return "Total"
        if i not in somar:
            return None
        # soma das parcelas como aparecem (arredondadas à precisão da coluna)
        return _soma(
            arredondar(v, colunas[i].tipo) for x in linhas if isinstance(v := x[i], Decimal)
        )

    return tuple(coluna(i) for i in range(len(colunas)))


# ------------------------------------------------------------------ 5. operações finalizadas
def _obs(m: Movimento, isento: bool) -> str:
    marcas: list[str] = []
    if m.categoria is Categoria.DAY_TRADE:
        marcas.append("DT")
    if "exercício" in m.regra:
        marcas.append("exercício")
    if "pó" in m.regra:
        marcas.append("pó")
    if m.isento_pme:
        marcas.append("isento PME")
    elif m.isentavel and isento:
        marcas.append("isento")
    if m.tipo is TipoAtivo.ETF_RF:
        marcas.append("trib. exclusiva")
    if m.tipo is TipoAtivo.FUNDO_INFRA:
        marcas.append("alíquota zero")
    if m.cv == "C" and m.custo_baixa:
        marcas.append("recompra")
    elif m.quantidade < 0:
        marcas.append("lançada")
    return ", ".join(marcas)
