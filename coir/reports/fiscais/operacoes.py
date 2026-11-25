"""Relatórios das operações: finalizadas por ativo, conferência do preço médio e livro das notas."""

from collections import defaultdict
from collections.abc import Iterable, Sequence
from datetime import date
from decimal import Decimal

from coir.domain.fiscal import (
    ZERO,
    Categoria,
    Movimento,
    TipoAtivo,
)
from coir.domain.nota import CategoriaCusto
from coir.domain.relatorio import DadosRelatorio
from coir.reports.agrupar import COLUNAS_POR_ATIVO, ItemOperacao, por_mes, total_por_ativo
from coir.reports.fiscais.comum import (
    CATEGORIA_CURTA,
    MESES,
    NOME_GRUPO,
    T,
    _cab,
    _cat,
    _dinheiro,
    _meses,
    _movimentos,
    _obs,
    _soma,
    _total,
)
from coir.reports.modelo import (
    Coluna,
    Relatorio,
    Secao,
    Valor,
)


def itens_finalizadas(
    colunas: Sequence[Coluna], linhas: Iterable[Sequence[Valor]]
) -> list[ItemOperacao]:
    """Linhas de operações finalizadas -> itens do agrupamento por mês e ativo."""
    nomes = [c.nome for c in colunas]

    def numero(x: Sequence[Valor], nome: str) -> Decimal | None:
        v = x[nomes.index(nome)]
        return v if isinstance(v, Decimal) else None

    def texto(x: Sequence[Valor], nome: str) -> str:
        return str(x[nomes.index(nome)] or "")

    saida: list[ItemOperacao] = []
    for x in linhas:
        dia = x[nomes.index("Data")]
        if isinstance(dia, date):
            saida.append(ItemOperacao(
                data=dia, ativo=texto(x, "Ativo"), corretora=texto(x, "Corretora"),
                categoria=texto(x, "Categoria"), sentido=texto(x, "C/V"),
                quantidade=numero(x, "Quantidade") or Decimal(0),
                abertura=numero(x, "Abertura"), encerramento=numero(x, "Encerramento"),
                resultado=numero(x, "Resultado") or Decimal(0), obs=texto(x, "Obs."),
            ))  # fmt: skip
    return saida


def _por_ativo_mes_a_mes(secoes: tuple[Secao, ...]) -> tuple[Secao, ...]:
    """Cada mês com uma linha por dia e ativo, para conferir com outro controle (e nos arquivos)."""
    if not secoes:
        return ()
    itens = itens_finalizadas(secoes[0].colunas, [x for s in secoes for x in s.linhas])
    return tuple(
        Secao(f"{ano} - {MESES[mes - 1]} - {mercado} - por ativo", COLUNAS_POR_ATIVO, tuple(linhas),
              total_por_ativo(linhas))
        for (ano, mes, mercado), linhas in por_mes(itens).items()
    )  # fmt: skip


def operacoes_finalizadas(d: DadosRelatorio) -> Relatorio:
    isento = {m.mes: _cat(m, Categoria.COMUM).isento > 0 for m in _meses(d)}
    colunas = (
        Coluna("Data", T.DATA),
        Coluna("Ativo"),
        Coluna("Corretora"),
        Coluna("Categoria"),
        Coluna("Tipo"),
        Coluna("C/V"),
        Coluna("Quantidade", T.QUANTIDADE),
        Coluna("Custo médio", T.PM),
        _dinheiro("Abertura"),
        _dinheiro("Encerramento"),
        _dinheiro("Resultado"),
        Coluna("Lucratividade", T.PERCENTUAL),
        Coluna("Obs."),
    )
    blocos: dict[tuple[int, str], list[tuple[Valor, ...]]] = defaultdict(list)
    for m in _movimentos(d):
        if not (m.resultado or m.custo_baixa):
            continue
        mercado = "BM&F" if m.tipo is TipoAtivo.FUTURO else "Bovespa"
        # posição vendida (lançada) que vira pó: quantidade e receita da abertura sem sinal
        quantidade, abertura = abs(m.quantidade), abs(m.custo_baixa)
        unit = abertura / quantidade if quantidade and abertura else None
        perc = m.resultado / abertura if abertura else None
        futuro = m.tipo is TipoAtivo.FUTURO
        blocos[(m.data.month, mercado)].append(
            (
                m.data,
                m.ticker,
                m.corretora or "",
                CATEGORIA_CURTA[m.categoria],
                NOME_GRUPO[m.grupo],
                {"V": "C/V", "C": "V/C"}.get(m.cv or "", ""),  # o que abriu / o que encerrou
                quantidade or None,
                unit,
                None if futuro else abertura,
                None if futuro else m.valor,
                m.resultado,
                perc,
                _obs(m, isento.get(m.data.month, False)),
            )
        )
    secoes = tuple(
        Secao(
            f"{d.ano} - {MESES[mes - 1]} - {mercado}",
            colunas,
            tuple(linhas),
            _total(linhas, colunas, {10}),
        )
        for (mes, mercado), linhas in sorted(
            blocos.items(), key=lambda x: (x[0][0], x[0][1] != "Bovespa")
        )
    )
    return Relatorio(
        "operacoes-finalizadas",
        "Operações finalizadas",
        _cab(d),
        (*secoes, *_por_ativo_mes_a_mes(secoes)),
        pdf=True,
        notas=(
            "Abertura: custo médio baixado (ou receita da venda, na recompra de "
            "posição vendida). Encerramento: alienação líquida (ou custo da "
            "recompra). Futuros: resultado dos ajustes no encerramento.",
        ),
    )


# ------------------------------------------------------------ 8. conferência de cálculo (PM)
def conferencia(d: DadosRelatorio) -> Relatorio:
    colunas = (
        Coluna("Data", T.DATA),
        Coluna("Corretora"),
        Coluna("Origem"),
        Coluna("Regra"),
        Coluna("C/V"),
        Coluna("Quantidade", T.QUANTIDADE),
        _dinheiro("Valor"),
        _dinheiro("Custos"),
        _dinheiro("Custo baixado"),
        _dinheiro("Resultado"),
        Coluna("Qtd. após", T.QUANTIDADE),
        Coluna("PM após", T.PM),
        _dinheiro("Custo após"),
    )
    anteriores: dict[str, Movimento] = {}
    trilhas: dict[str, list[Movimento]] = defaultdict(list)
    for m in d.apuracao.movimentos:
        if m.titular != d.titular or m.ativo is None or m.data.year > d.ano:
            continue
        if m.data.year < d.ano:
            anteriores[m.ticker] = m
        else:
            trilhas[m.ticker].append(m)
    secoes: list[Secao] = []
    for ticker in sorted(trilhas):
        linhas: list[tuple[Valor, ...]] = []
        if (a := anteriores.get(ticker)) is not None and a.quantidade_apos:
            linhas.append(
                (
                    None,
                    "",
                    "saldo inicial",
                    f"posição em 31/12/{d.ano - 1}",
                    "",
                    None,
                    None,
                    None,
                    None,
                    None,
                    a.quantidade_apos,
                    a.pm_apos,
                    a.custo_apos,
                )
            )
        for m in trilhas[ticker]:
            linhas.append(
                (
                    m.data,
                    m.corretora or "",
                    m.origem,
                    m.regra,
                    m.cv or "",
                    m.quantidade,
                    m.valor,
                    m.custos,
                    m.custo_baixa,
                    m.resultado,
                    m.quantidade_apos,
                    m.pm_apos,
                    m.custo_apos,
                )
            )
        secoes.append(Secao(ticker, colunas, tuple(linhas), _total(linhas, colunas, {9})))
    return Relatorio(
        "conferencia-pm",
        "Conferência de cálculo (custo médio)",
        _cab(d),
        tuple(secoes),
        notas=("PM fiscal por titular (todas as corretoras e contas; IN RFB 1.585/2015 art. 58).",),
    )


# --------------------------------------------------------------------------- 10. livro de notas
CUSTOS = tuple(CategoriaCusto)


def livro_notas(d: DadosRelatorio) -> Relatorio:
    colunas = (
        Coluna("Data", T.DATA),
        Coluna("Corretora"),
        Coluna("Conta"),
        Coluna("Número"),
        Coluna("Segmento"),
        _dinheiro("Valor das operações"),
        _dinheiro("Custos"),
        _dinheiro("IRRF sobre operações"),
        _dinheiro("IRRF day trade"),
        _dinheiro("Líquido"),
        *(_dinheiro(c.value.replace("_", " ").capitalize()) for c in CUSTOS),
    )
    notas: list[tuple[Valor, ...]] = []
    ops: list[tuple[Valor, ...]] = []
    bmf: list[tuple[Valor, ...]] = []
    for item in d.notas:
        n = item.nota
        por_cat: dict[CategoriaCusto, Decimal] = defaultdict(Decimal)
        if n.resumo_bovespa is not None:
            r = n.resumo_bovespa
            for c in r.custos:
                por_cat[c.categoria] += c.valor
            valor, irrf, irrf_dt = r.valor_operacoes, r.irrf, r.irrf_day_trade
            liquido = r.liquido.valor if r.liquido.dc != "D" else -r.liquido.valor
        elif n.resumo_bmf is not None:
            r2 = n.resumo_bmf
            for c in r2.custos:
                por_cat[c.categoria] += c.valor
            valor, irrf, irrf_dt = r2.valor_negocios.valor, r2.irrf, r2.irrf_day_trade
            liquido = r2.liquido.valor if r2.liquido.dc != "D" else -r2.liquido.valor
        else:
            valor = irrf = irrf_dt = liquido = ZERO
        notas.append(
            (
                n.data_pregao,
                n.corretora,
                n.conta,
                n.numero,
                n.segmento.value,
                valor,
                _soma(por_cat.values()),
                irrf,
                irrf_dt,
                liquido,
                *(por_cat.get(c) for c in CUSTOS),
            )
        )
        for o in item.operacoes:
            ops.append(
                (
                    o.data,
                    o.corretora,
                    o.nota,
                    o.papel,
                    o.mercado,
                    o.cv,
                    "DT" if o.categoria == "DAY_TRADE" else "",
                    o.quantidade,
                    o.preco,
                    o.total,
                )
            )
        for b in n.negocios_bmf:
            bmf.append(
                (
                    n.data_pregao,
                    n.corretora,
                    n.numero,
                    b.mercadoria,
                    b.tipo,
                    b.cv,
                    b.quantidade,
                    b.preco,
                    b.valor if b.dc != "D" else -b.valor,
                )
            )
    secao_notas = Secao(
        "Notas", colunas, tuple(notas), _total(notas, colunas, {5, 6, 7, 8, 9}) if notas else None
    )
    colunas_ops = (
        Coluna("Data", T.DATA),
        Coluna("Corretora"),
        Coluna("Nota"),
        Coluna("Ativo"),
        Coluna("Mercado"),
        Coluna("C/V"),
        Coluna("Obs."),
        Coluna("Quantidade", T.QUANTIDADE),
        Coluna("Preço", T.PM),
        _dinheiro("Total"),
    )
    colunas_bmf = (
        Coluna("Data", T.DATA),
        Coluna("Corretora"),
        Coluna("Nota"),
        Coluna("Mercadoria"),
        Coluna("Tipo"),
        Coluna("C/V"),
        Coluna("Quantidade", T.QUANTIDADE),
        Coluna("Preço", T.PM),
        _dinheiro("Ajuste"),
    )
    return Relatorio(
        "livro-de-notas",
        "Livro de notas",
        _cab(d),
        (
            secao_notas,
            Secao("Operações (Bovespa)", colunas_ops, tuple(ops)),
            Secao("Negócios BM&F", colunas_bmf, tuple(bmf)),
        ),
    )
