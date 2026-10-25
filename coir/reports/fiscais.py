"""Relatórios fiscais de um titular e um ano (F5), montados só a partir de `DadosRelatorio`.

Cada função devolve um `Relatorio` neutro: a exportação (xlsx, csv, pdf) e a interface (F6)
usam o mesmo objeto. Nenhum cálculo fiscal novo aqui: os valores vêm da apuração do motor
(regra 12: cada número tem origem no motor ou na nota).
"""

from collections import defaultdict
from collections.abc import Callable, Iterable
from dataclasses import replace
from decimal import Decimal

from coir.domain.fiscal import (
    ZERO,
    Categoria,
    Ficha,
    FichaMes,
    Grupo,
    MesApurado,
    Movimento,
    PosicaoAtivo,
    ResultadoCategoria,
    TipoAtivo,
    Visao,
)
from coir.domain.nota import CategoriaCusto
from coir.domain.proventos import Provento, TipoProvento
from coir.domain.relatorio import DadosRelatorio
from coir.reports.irpf import declaracao
from coir.reports.modelo import (
    Cabecalho,
    Coluna,
    Relatorio,
    Secao,
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
        or r.vendas_acoes
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


# ------------------------------------------------------------------------- 1. resumo fiscal
def resumo_fiscal(d: DadosRelatorio) -> Relatorio:
    meses = _meses(d)
    secoes: list[Secao] = []
    for cat in Categoria:
        comum = cat is Categoria.COMUM
        colunas = (
            Coluna("Mês"),
            *(_dinheiro(NOME_GRUPO[g]) for g in GRUPOS),
            _dinheiro("Resultado"),
            *((_dinheiro("Vendas de ações"), _dinheiro("Isento")) if comum else ()),
            _dinheiro("Prejuízo anterior"),
            _dinheiro("Base de cálculo"),
            Coluna("Alíquota", T.PERCENTUAL),
            _dinheiro("Imposto"),
            _dinheiro("Prejuízo a compensar"),
        )
        linhas: list[tuple[Valor, ...]] = []
        for m in meses if _ativa(meses, cat) else []:
            r = _cat(m, cat)
            linhas.append(
                (
                    MESES[m.mes - 1],
                    *(r.por_grupo.get(g, ZERO) for g in GRUPOS),
                    r.resultado,
                    *((r.vendas_acoes, r.isento) if comum else ()),
                    r.prejuizo_anterior,
                    r.base,
                    r.aliquota,
                    r.imposto,
                    r.prejuizo_a_compensar,
                )
            )
        n = len(GRUPOS) + 1  # colunas dos grupos + resultado
        somar = set(range(1, n + 1)) | ({n + 1, n + 2} if comum else set()) | {len(colunas) - 2}
        secoes.append(
            Secao(
                NOME_CATEGORIA[cat],
                colunas,
                tuple(linhas),
                _total(linhas, colunas, somar) if linhas else None,
            )
        )
    colunas = (
        Coluna("Mês"),
        _dinheiro("Imposto"),
        _dinheiro("IRRF do mês"),
        _dinheiro("IRRF saldo anterior"),
        _dinheiro("IRRF deduzido"),
        _dinheiro("IRRF saldo"),
        _dinheiro("Imposto devido"),
        _dinheiro("DARF acumulado anterior"),
        _dinheiro("DARF"),
        _dinheiro("Acumulado para o mês seguinte"),
        Coluna("Vencimento", T.DATA),
        _dinheiro("Isento PME"),
    )
    linhas = [
        (
            MESES[m.mes - 1],
            m.imposto,
            m.irrf_mes,
            m.irrf_saldo_anterior,
            m.irrf_deduzido,
            m.irrf_saldo,
            m.imposto_devido,
            m.darf_acumulado_anterior,
            m.darf,
            m.darf_acumulado,
            m.vencimento,
            m.isento_pme,
        )
        for m in meses
    ]
    secoes.append(
        Secao(
            "Consolidação do mês",
            colunas,
            tuple(linhas),
            _total(linhas, colunas, {1, 2, 4, 6, 8, 11}) if linhas else None,
            nota="IRRF sobre vendas comuns (Lei 11.033/2004) e sobre ganho de day trade, com "
            "saldo no ano; DARF abaixo do valor mínimo acumula (Lei 9.430/1996 art. 68).",
        )
    )
    return Relatorio("resumo-fiscal", "Resumo fiscal do ano", _cab(d), tuple(secoes), pdf=True)


# ------------------------------------------------------------------- 2. DARFs e meses operados
def darfs(d: DadosRelatorio) -> Relatorio:
    meses = _meses(d)
    colunas = (
        Coluna("Competência"),
        Coluna("Código"),
        _dinheiro("Valor"),
        Coluna("Vencimento", T.DATA),
    )
    linhas: list[tuple[Valor, ...]] = [
        (f"{m.mes:02d}/{m.ano}", d.codigo_darf, m.darf, m.vencimento) for m in meses if m.darf
    ]
    pagar = Secao(
        f"DARFs a pagar (código {d.codigo_darf})",
        colunas,
        tuple(linhas),
        _total(linhas, colunas, {2}) if linhas else None,
        nota="Vencimento no último dia útil do mês seguinte (RIR/2018 art. 915).",
    )
    notas = defaultdict[int, int](int)
    for n in d.notas:
        notas[n.nota.data_pregao.month] += 1
    por_mes = {m.mes: m for m in meses}
    colunas = (
        Coluna("Mês"),
        Coluna("Notas", T.INTEIRO),
        _dinheiro("Imposto devido"),
        _dinheiro("DARF"),
        _dinheiro("Acumulado para o mês seguinte"),
    )
    linhas = []
    for mes in range(1, 13):
        m = por_mes.get(mes)
        if m is None and not notas[mes]:
            continue
        linhas.append(
            (
                MESES[mes - 1],
                notas[mes],
                m.imposto_devido if m else None,
                m.darf if m else None,
                m.darf_acumulado if m else None,
            )
        )
    total = (
        "Total",
        sum(notas.values()),
        _soma(m.imposto_devido for m in meses),
        _soma(m.darf for m in meses),
        None,
    )
    operados = Secao("Meses operados", colunas, tuple(linhas), total if linhas else None)
    return Relatorio("darfs", "DARFs do ano e meses operados", _cab(d), (pagar, operados), pdf=True)


# ------------------------------------------------------------- 3. ordem do programa da Receita
def irpf_mensal(d: DadosRelatorio) -> Relatorio:
    meses = _meses(d)
    comum, dt = Categoria.COMUM, Categoria.DAY_TRADE
    colunas = (
        Coluna("Mês"),
        *(_dinheiro(f"Comum: {NOME_GRUPO[g]}") for g in GRUPOS),
        *(_dinheiro(f"Day trade: {NOME_GRUPO[g]}") for g in GRUPOS),
    )
    linhas: list[tuple[Valor, ...]] = [
        (
            MESES[m.mes - 1],
            *(_cat(m, comum).por_grupo.get(g, ZERO) for g in GRUPOS),
            *(_cat(m, dt).por_grupo.get(g, ZERO) for g in GRUPOS),
        )
        for m in meses
    ]
    mercados = Secao(
        "Resultados por mercado",
        colunas,
        tuple(linhas),
        _total(linhas, colunas, set(range(1, 2 * len(GRUPOS) + 1))) if linhas else None,
        nota="Termo: a venda à vista na data da liquidação vai para a linha do termo (IN RFB "
        "1.585/2015 art. 62 I); vendido depois, à vista. Ouro não é operado.",
    )

    def bloco(r: ResultadoCategoria) -> tuple[Valor, ...]:
        return (
            r.resultado,
            r.prejuizo_anterior,
            r.base,
            r.prejuizo_a_compensar,
            r.aliquota,
            r.imposto,
        )

    nomes = (
        "resultado líquido",
        "resultado negativo até o mês anterior",
        "base de cálculo",
        "prejuízo a compensar",
        "alíquota",
        "imposto devido",
    )
    colunas_r = (
        Coluna("Mês"),
        *(Coluna(n.capitalize(), T.PERCENTUAL if n == "alíquota" else T.DINHEIRO) for n in nomes),
    )
    resultados = [
        Secao(f"Resultados - {titulo}", colunas_r,
              tuple((MESES[m.mes - 1], *bloco(_cat(m, cat))) for m in meses))
        for titulo, cat in (("operações comuns", comum), ("day trade", dt))
    ]  # fmt: skip
    colunas = (
        Coluna("Mês"),
        _dinheiro("Total do imposto devido"),
        _dinheiro("IR fonte de day trade no mês"),
        _dinheiro("IR fonte de day trade meses anteriores"),
        _dinheiro("IR fonte de day trade a compensar"),
        _dinheiro("IR fonte (Lei 11.033/2004) no mês"),
        _dinheiro("IR fonte (Lei 11.033/2004) meses anteriores"),
        _dinheiro("IR fonte (Lei 11.033/2004) a compensar"),
        _dinheiro("Imposto a pagar"),
        _dinheiro("Imposto pago (calculado)"),
    )

    def ficha(f: FichaMes) -> tuple[Valor, ...]:
        return (f.imposto, f.irrf_day_trade_mes, f.irrf_day_trade_anterior,
                f.irrf_day_trade_a_compensar, f.irrf_mes, f.irrf_anterior, f.irrf_a_compensar,
                f.a_pagar, f.pago)  # fmt: skip

    linhas = [(MESES[m.mes - 1], *ficha(m.ficha(Ficha.COMUNS))) for m in meses]
    consolidacao = Secao(
        "Consolidação do mês",
        colunas,
        tuple(linhas),
        nota="Só comuns e day trade (FII tem ficha própria). A DARF deduz o IRRF de um saldo "
        "único; cada ficha recebe o IRRF que consumiu nele (o de day trade primeiro) e o retido "
        "nela que sobrou, sem usar o mesmo valor duas vezes. Imposto pago: o a pagar do mês "
        "quando a DARF que o cobriu (inclusive acumulada abaixo de R$ 10) já foi emitida.",
    )
    colunas = (
        Coluna("Mês"),
        _dinheiro("Total em alienações de ações (Lei 11.033/2004)"),
        _dinheiro("Lucro isento (Lei 11.033/2004)"),
        _dinheiro("Resultado isento de PME (Lei 13.043/2014 art. 16)"),
    )
    linhas = [
        (MESES[m.mes - 1], _cat(m, comum).vendas_acoes, _cat(m, comum).isento, m.isento_pme)
        for m in meses
    ]
    outras = Secao(
        "Outras informações (só para consulta)",
        colunas,
        tuple(linhas),
        _total(linhas, colunas, {1, 2, 3}) if linhas else None,
    )
    colunas = (
        Coluna("Mês"),
        *(Coluna(n.capitalize(), T.PERCENTUAL if n == "alíquota" else T.DINHEIRO) for n in nomes),
    )
    colunas = (
        *colunas,
        _dinheiro("IR fonte (Lei 11.033/2004) no mês"),
        _dinheiro("IR fonte (Lei 11.033/2004) meses anteriores"),
        _dinheiro("IR fonte (Lei 11.033/2004) a compensar"),
        _dinheiro("Imposto a pagar"),
        _dinheiro("Imposto pago (calculado)"),
    )
    fiis = meses if _ativa(meses, Categoria.FII) or any(
        m.ficha(Ficha.FII).irrf_mes for m in meses) else []  # fmt: skip
    linhas = []
    for m in fiis:
        f = m.ficha(Ficha.FII)
        linhas.append((MESES[m.mes - 1], *bloco(_cat(m, Categoria.FII)), f.irrf_mes,
                       f.irrf_anterior, f.irrf_a_compensar, f.a_pagar, f.pago))  # fmt: skip
    usado_dt = [m for m in fiis if m.ficha(Ficha.FII).irrf_day_trade_mes]
    fii = Secao(
        "Fundos imobiliários e Fiagro",
        colunas,
        tuple(linhas),
        nota="IR fonte: o 0,005% retido em vendas de FII que não abateu a ficha de comuns, mais "
        "o de comuns que abateu esta ficha."
        + (" Houve IRRF de day trade deduzido do imposto de FII (sem linha no programa): "
           + ", ".join(MESES[m.mes - 1] for m in usado_dt) + "." if usado_dt else ""),
    )  # fmt: skip
    return Relatorio(
        "irpf-mensal",
        "Renda variável mês a mês (ordem do programa IRPF)",
        _cab(d),
        (mercados, *resultados, consolidacao, outras, fii),
        pdf=True,
        notas=("Mesma ordem dos blocos do programa da Receita; as fichas por ano ficam na F8.",),
    )


# --------------------------------------------------------- 4. IRRF anual e tributação exclusiva
def irrf_exclusiva(d: DadosRelatorio) -> Relatorio:
    anual = [a for a in d.apuracao.irrf_anual if a.titular == d.titular and a.ano == d.ano]
    colunas = (
        Coluna("Ano", T.INTEIRO),
        _dinheiro("IRRF de vendas comuns retido"),
        _dinheiro("IRRF day trade retido"),
        _dinheiro("Saldo de vendas comuns (declaração)"),
        _dinheiro("Na ficha de comuns"),
        _dinheiro("Na ficha de FII"),
        _dinheiro("Saldo day trade (restituição)"),
    )
    irrf = Secao(
        "IRRF do ano",
        colunas,
        tuple(
            (
                a.ano,
                a.comum,
                a.day_trade,
                a.saldo_comum,
                a.saldo_comum_fichas.get(Ficha.COMUNS, ZERO),
                a.saldo_comum_fichas.get(Ficha.FII, ZERO),
                a.saldo_day_trade,
            )
            for a in anual
        ),
        nota=(
            "O saldo de dezembro não passa para janeiro: o de vendas comuns vai para a "
            "declaração anual e o de day trade só por restituição (IN RFB 1.585/2015 "
            "arts. 63 e 65)."
        )
        + ("" if anual else " Ano em curso: o saldo só fecha em dezembro."),
    )
    exclusivas = [
        e for e in d.apuracao.exclusiva if e.titular == d.titular and e.data.year == d.ano
    ]
    colunas = (
        Coluna("Data", T.DATA),
        Coluna("Ativo"),
        Coluna("Quantidade", T.QUANTIDADE),
        _dinheiro("Alienação"),
        _dinheiro("Custo"),
        _dinheiro("Resultado"),
    )
    linhas: list[tuple[Valor, ...]] = [
        (e.data, e.ticker, e.quantidade, e.alienacao, e.custo, e.resultado) for e in exclusivas
    ]
    exclusiva = Secao(
        "Tributação exclusiva (ETF de renda fixa)",
        colunas,
        tuple(linhas),
        _total(linhas, colunas, {3, 4, 5}) if linhas else None,
        nota="Lei 13.043/2014 art. 2º: imposto retido na fonte; fora da apuração mensal e do "
        "DARF. Resultado estimado pelo custo médio do COIR: confira com o informe de "
        "rendimentos da instituição.",
    )
    return Relatorio(
        "irrf-e-exclusiva",
        "IRRF do ano e tributação exclusiva",
        _cab(d),
        (irrf, exclusiva),
        pdf=True,
    )


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
    if m.cv == "C" and m.custo_baixa:
        marcas.append("recompra")
    elif m.quantidade < 0:
        marcas.append("lançada")
    return ", ".join(marcas)


def operacoes_finalizadas(d: DadosRelatorio) -> Relatorio:
    isento = {m.mes: _cat(m, Categoria.COMUM).isento > 0 for m in _meses(d)}
    colunas = (
        Coluna("Data", T.DATA),
        Coluna("Ativo"),
        Coluna("Corretora"),
        Coluna("Categoria"),
        Coluna("Tipo"),
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
            _total(linhas, colunas, {9}),
        )
        for (mes, mercado), linhas in sorted(
            blocos.items(), key=lambda x: (x[0][0], x[0][1] != "Bovespa")
        )
    )
    return Relatorio(
        "operacoes-finalizadas",
        "Operações finalizadas",
        _cab(d),
        secoes,
        pdf=True,
        notas=(
            "Abertura: custo médio baixado (ou receita da venda, na recompra de "
            "posição vendida). Encerramento: alienação líquida (ou custo da "
            "recompra). Futuros: resultado dos ajustes no encerramento.",
        ),
    )


# ---------------------------------------------------------------- 6 e 7. evolução e prejuízos
def evolucao(d: DadosRelatorio) -> Relatorio:
    colunas = (
        Coluna("Mês"),
        *(_dinheiro(NOME_CATEGORIA[c]) for c in Categoria),
        _dinheiro("Isento (ações e PME)"),
        _dinheiro("Total do mês"),
        _dinheiro("Acumulado no ano"),
    )
    linhas: list[tuple[Valor, ...]] = []
    acumulado = ZERO
    for m in _meses(d):
        cats = [_cat(m, c).resultado for c in Categoria]
        isento = _cat(m, Categoria.COMUM).isento + m.isento_pme
        total = _soma(cats) + isento
        acumulado += total
        linhas.append((MESES[m.mes - 1], *cats, isento, total, acumulado))
    secao = Secao(
        "Resultado mês a mês",
        colunas,
        tuple(linhas),
        _total(linhas, colunas, {1, 2, 3, 4, 5}) if linhas else None,
    )
    return Relatorio("evolucao", "Evolução do resultado", _cab(d), (secao,))


def prejuizos(d: DadosRelatorio) -> Relatorio:
    colunas = (Coluna("Mês"), *(_dinheiro(NOME_CATEGORIA[c]) for c in Categoria))
    linhas = tuple(
        (MESES[m.mes - 1], *(_cat(m, c).prejuizo_a_compensar for c in Categoria)) for m in _meses(d)
    )
    secao = Secao(
        "Prejuízo a compensar no fim do mês",
        colunas,
        linhas,
        nota="Comum só com comum, day trade só com day trade, FII só com FII; sem "
        "prazo (IN RFB 1.585/2015).",
    )
    return Relatorio("prejuizos", "Prejuízos a compensar", _cab(d), (secao,))


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


# ------------------------------------------------------------- 9. carteira e bens e direitos
BLOCOS_CARTEIRA: tuple[tuple[str, Callable[[PosicaoAtivo], bool]], ...] = (
    (
        "Ações, BDRs, ETFs e units",
        lambda p: (
            p.quantidade > 0
            and p.tipo
            in (TipoAtivo.ACAO, TipoAtivo.UNIT, TipoAtivo.BDR, TipoAtivo.ETF, TipoAtivo.ETF_RF)
        ),
    ),
    ("Fundos imobiliários e Fiagro", lambda p: p.quantidade > 0 and p.tipo is TipoAtivo.FII),
    ("Opções", lambda p: p.quantidade > 0 and p.tipo is TipoAtivo.OPCAO),
    ("Posições vendidas", lambda p: p.quantidade < 0),
)


def _onde(p: PosicaoAtivo, contas: dict[int, str]) -> tuple[str, str]:
    return p.corretora or "", contas.get(p.conta, str(p.conta)) if p.conta is not None else ""


def carteira(d: DadosRelatorio) -> Relatorio:
    colunas = (
        Coluna("Ativo"),
        Coluna("Corretora"),
        Coluna("Conta"),
        Coluna("Vencimento", T.DATA),
        Coluna("Quantidade", T.QUANTIDADE),
        Coluna("Custo médio", T.PM),
        _dinheiro("Custo total"),
    )
    secoes: list[Secao] = []
    for titulo, filtro in BLOCOS_CARTEIRA:
        posicoes = sorted(
            (p for p in d.carteira if p.titular == d.titular and filtro(p)),
            key=lambda p: (p.ticker, p.corretora or "", p.conta or 0),
        )
        linhas: list[tuple[Valor, ...]] = [
            (p.ticker, *_onde(p, d.contas), p.vencimento, p.quantidade, p.pm, p.custo)
            for p in posicoes
        ]
        secoes.append(
            Secao(titulo, colunas, tuple(linhas), _total(linhas, colunas, {6}) if linhas else None)
        )
    if d.visao is Visao.TITULAR:
        visao = "PM fiscal por titular (apuração e IRPF)"
    else:
        visao = f"visão operacional por {d.visao.value} (não é o PM do imposto)"
    return Relatorio(
        "carteira",
        "Carteira",
        _cab(d, f"Posição em {d.data_carteira:%d/%m/%Y}"),
        tuple(secoes),
        pdf=True,
        notas=(f"Custo médio: {visao}.",),
    )


def bens_e_direitos(d: DadosRelatorio) -> Relatorio:
    colunas = (
        Coluna("Ativo"),
        Coluna("Quantidade", T.QUANTIDADE),
        Coluna("Custo médio", T.PM),
        _dinheiro("Custo total"),
    )
    secoes: list[Secao] = []
    blocos = (
        (
            "Ações, ETFs, BDRs e units",
            lambda p: p.tipo not in (TipoAtivo.OPCAO, TipoAtivo.FUTURO, TipoAtivo.FII),
        ),
        ("Fundos imobiliários e Fiagro", lambda p: p.tipo is TipoAtivo.FII),
        ("Opções", lambda p: p.tipo is TipoAtivo.OPCAO),
    )
    for titulo, filtro in blocos:
        posicoes = sorted(
            (
                p
                for p in d.carteira_fiscal
                if p.titular == d.titular and p.quantidade > 0 and filtro(p)
            ),
            key=lambda p: p.ticker,
        )
        linhas: list[tuple[Valor, ...]] = [
            (p.ticker, p.quantidade, p.pm, p.custo) for p in posicoes
        ]
        secoes.append(
            Secao(titulo, colunas, tuple(linhas), _total(linhas, colunas, {3}) if linhas else None)
        )
    return Relatorio(
        "bens-e-direitos",
        "Bens e Direitos (custo de aquisição)",
        _cab(d, f"Situação em {d.ate:%d/%m/%Y}"),
        tuple(secoes),
        pdf=True,
        notas=(
            "Custo médio fiscal por titular, somando todas as corretoras. "
            "Posições vendidas não entram em Bens e Direitos.",
        ),
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


# --------------------------------------------------------------------------------- 12. proventos
NOME_PROVENTO = {
    TipoProvento.DIVIDENDO: "Dividendo", TipoProvento.JCP: "JCP",
    TipoProvento.RENDIMENTO_FII: "Rendimento de FII", TipoProvento.REEMBOLSO_BTC: "Reembolso BTC",
    TipoProvento.ATUALIZACAO: "Atualização (Selic)",
    TipoProvento.RESTITUICAO_CAPITAL: "Restituição de capital",
    TipoProvento.FORA_DO_ESCOPO: "Fora da v3.0",
}  # fmt: skip


def _por_empresa(lista: list[Provento], jcp: bool = False) -> Secao | None:
    grupos: dict[tuple[str, str], list[Provento]] = defaultdict(list)
    for p in lista:
        grupos[(p.codigo[:4], p.empresa)].append(p)
    if jcp:
        colunas = (Coluna("Empresa"), Coluna("Códigos"), _dinheiro("Líquido (código 10)"),
                   _dinheiro("Bruto (estimado)"), _dinheiro("IRRF (estimado)"))  # fmt: skip
    else:
        colunas = (Coluna("Empresa"), Coluna("Códigos"), _dinheiro("Valor"))
    linhas: list[tuple[Valor, ...]] = []
    for (_, empresa), ps in sorted(grupos.items()):
        codigos = ", ".join(sorted({p.codigo for p in ps}))
        liquido = _soma(p.liquido for p in ps)
        if jcp:
            linhas.append((empresa, codigos, liquido, _soma(p.bruto for p in ps),
                           _soma(p.irrf for p in ps)))  # fmt: skip
        else:
            linhas.append((empresa, codigos, liquido))
    if not linhas:
        return None
    return Secao("", colunas, tuple(linhas), _total(linhas, colunas, set(range(2, len(colunas)))))


def proventos(d: DadosRelatorio) -> Relatorio:
    """Proventos do ano para o IRPF (declarados no ano: o JCP creditado no ano e pago no seguinte
    entra no ano do crédito, decisão do usuário de 19/10/2026)."""
    do_ano = [p for p in d.proventos if p.ano_declaracao == d.ano]
    secoes: list[Secao] = []
    blocos = (
        (TipoProvento.DIVIDENDO, "Dividendos (rendimentos isentos, código 09)", False,
         "Lei 9.249/1995 art. 10. A partir de 2026, veja os avisos da Lei 15.270/2025."),
        (TipoProvento.JCP, "Juros sobre capital próprio (tributação exclusiva, código 10)", True,
         "Valor líquido creditado; bruto e IRRF estimados pela alíquota retida (15% até 2025, "
         "17,5% desde 2026, LC 224/2025): vale o informe da empresa."),
        (TipoProvento.RENDIMENTO_FII, "Rendimentos de FII e Fiagro (isentos)", False,
         "Lei 11.033/2004 art. 3º III: isentos para pessoa física se o fundo cumpre os "
         "requisitos."),
        (TipoProvento.ATUALIZACAO, "Atualização de proventos pela Selic (tributação exclusiva)",
         True, "Lei 8.981/1995 art. 65 §4º c; IRRF retido pela empresa (renda fixa). Código 06 ou "
         "12 na ficha exclusiva, conforme o informe da empresa."),
        (TipoProvento.REEMBOLSO_BTC, "Reembolsos de proventos de ações emprestadas (BTC)", False,
         "IN RFB 1.585/2015 art. 74 até 2025; desde 2026 (Lei 15.265/2025) isento por equivalência "
         "para dividendos (lei omissa)."),
    )  # fmt: skip
    for tipo, titulo, jcp, nota in blocos:
        secao = _por_empresa([p for p in do_ano if p.tipo is tipo], jcp)
        if secao is not None:
            secoes.append(replace(secao, titulo=titulo, nota=nota))
    a_receber = [p for p in do_ano if p.a_receber_em_31_12]
    if a_receber:
        colunas = (Coluna("Empresa"), Coluna("Código"), Coluna("Pago em", T.DATA),
                   _dinheiro("Líquido a receber"))  # fmt: skip
        linhas: list[tuple[Valor, ...]] = [
            (p.empresa, p.codigo, p.data, p.liquido) for p in a_receber
        ]
        secoes.append(Secao(f"JCP creditado e a receber em 31/12/{d.ano} (Bens e Direitos 99.07)",
                            colunas, tuple(linhas), _total(linhas, colunas, {3})))  # fmt: skip
    colunas = (Coluna("Data", T.DATA), Coluna("Código"), Coluna("Tipo"), Coluna("Instituição"),
               Coluna("Quantidade", T.QUANTIDADE), _dinheiro("Líquido"), _dinheiro("Bruto"),
               _dinheiro("IRRF"), Coluna("Observação"))  # fmt: skip
    ordenados = sorted(do_ano, key=lambda p: (p.data, p.codigo))
    todos: list[tuple[Valor, ...]] = [
        (p.data, p.codigo, NOME_PROVENTO[p.tipo], p.instituicao, p.quantidade, p.liquido, p.bruto,
         p.irrf, p.observacao) for p in ordenados
    ]  # fmt: skip
    secoes.append(Secao("Lançamentos do ano (movimentação da B3)", colunas, tuple(todos)))
    avisos = tuple(a for a in d.avisos_proventos if f"{d.ano}" in a[:10])
    return Relatorio(
        "proventos", "Proventos recebidos", _cab(d), tuple(secoes), pdf=True,
        notas=(*avisos, "Fonte: movimentação da Área do Investidor da B3. O CNPJ da fonte "
               "pagadora está no informe de rendimentos da empresa."),
    )  # fmt: skip


TODOS: dict[str, Callable[[DadosRelatorio], Relatorio]] = {
    "resumo-fiscal": resumo_fiscal,
    "darfs": darfs,
    "irpf-mensal": irpf_mensal,
    "irrf-e-exclusiva": irrf_exclusiva,
    "operacoes-finalizadas": operacoes_finalizadas,
    "evolucao": evolucao,
    "prejuizos": prejuizos,
    "conferencia-pm": conferencia,
    "carteira": carteira,
    "bens-e-direitos": bens_e_direitos,
    "livro-de-notas": livro_notas,
    "proventos": proventos,
    "declaracao-irpf": declaracao,
}


ABAS = {
    "resumo-fiscal": "Resumo fiscal", "darfs": "DARFs", "irpf-mensal": "IRPF mês a mês",
    "irrf-e-exclusiva": "IRRF e exclusiva", "operacoes-finalizadas": "Operações finalizadas",
    "evolucao": "Evolução", "prejuizos": "Prejuízos", "conferencia-pm": "Conferência do PM",
    "carteira": "Carteira", "bens-e-direitos": "Bens e Direitos",
    "livro-de-notas": "Livro de notas", "proventos": "Proventos",
    "declaracao-irpf": "Declaração IRPF",
}  # fmt: skip


def relatorios(d: DadosRelatorio, chaves: Iterable[str] | None = None) -> list[Relatorio]:
    escolhidos = list(TODOS) if chaves is None else list(chaves)
    return [replace(TODOS[c](d), aba=ABAS[c]) for c in escolhidos]
