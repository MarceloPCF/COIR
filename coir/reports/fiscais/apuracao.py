"""Relatórios da apuração mensal: resumo fiscal, DARF, IRPF mês a mês, IRRF e tributação exclusiva,
evolução e prejuízos a compensar."""

from collections import defaultdict

from coir.domain.fiscal import (
    ZERO,
    Categoria,
    Ficha,
    FichaMes,
    ResultadoCategoria,
    TipoAtivo,
)
from coir.domain.relatorio import DadosRelatorio
from coir.reports.fiscais.comum import (
    GRUPOS,
    MESES,
    NOME_CATEGORIA,
    NOME_GRUPO,
    T,
    _ativa,
    _cab,
    _cat,
    _dinheiro,
    _meses,
    _soma,
    _total,
)
from coir.reports.modelo import (
    Coluna,
    Relatorio,
    Secao,
    Valor,
)


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
            *(
                (_dinheiro("Vendas de ações (soma do limite)"), _dinheiro("Isento"))
                if comum
                else ()
            ),
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
                    *((r.vendas_limite, r.isento) if comum else ()),
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
        _dinheiro("Perda PME (não compensa)"),
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
            m.perda_pme,
        )
        for m in meses
    ]
    secoes.append(
        Secao(
            "Imposto, IRRF e DARF do mês (como a DARF é calculada)",
            colunas,
            tuple(linhas),
            _total(linhas, colunas, {1, 2, 4, 6, 8, 11, 12}) if linhas else None,
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
    colunas = (Coluna("Mês"), *(_dinheiro(NOME_GRUPO[g]) for g in GRUPOS))
    nota_termo = (
        "Termo: a venda à vista na data da liquidação vai para a linha do termo (IN RFB "
        "1.585/2015 art. 62 I); vendido depois, à vista. Ouro não é operado."
    )

    def por_mercado(titulo: str, cats: tuple[Categoria, ...], nota: str | None) -> Secao:
        linhas: list[tuple[Valor, ...]] = [
            (MESES[m.mes - 1],
             *(_soma(_cat(m, c).por_grupo.get(g, ZERO) for c in cats) for g in GRUPOS))
            for m in meses
        ]  # fmt: skip
        total = _total(linhas, colunas, set(range(1, len(GRUPOS) + 1))) if linhas else None
        return Secao(titulo, colunas, tuple(linhas), total, nota=nota or "")

    mercados = [
        por_mercado("Resultados por mercado - operações comuns", (comum,), None),
        por_mercado("Resultados por mercado - day trade", (dt,), None),
        por_mercado("Resultados por mercado - comuns e day trade somados (só para conferir)",
                    (comum, dt), nota_termo),
    ]  # fmt: skip

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
        "Consolidação do mês (quadro final da ficha de renda variável do programa IRPF)",
        colunas,
        tuple(linhas),
        nota='Os valores a copiar no quadro "Consolidação do mês" do programa. Só comuns e '
        "day trade (FII tem ficha própria). A DARF deduz o IRRF de um saldo "
        "único; cada ficha recebe o IRRF que consumiu nele (o de day trade primeiro) e o retido "
        "nela que sobrou, sem usar o mesmo valor duas vezes. Imposto pago: o a pagar do mês "
        "quando a DARF que o cobriu (inclusive acumulada abaixo de R$ 10) já foi emitida.",
    )
    colunas = (
        Coluna("Mês"),
        _dinheiro("Total em alienações de ações, com day trade (Lei 11.033/2004)"),
        _dinheiro("Lucro isento (Lei 11.033/2004)"),
        _dinheiro("Lucro isento de PME (Lei 13.043/2014 art. 16)"),
        _dinheiro("Perda em ação de PME (não compensa: IN RFB 1.585/2015 art. 67 §2º)"),
    )
    linhas: list[tuple[Valor, ...]] = [
        (MESES[m.mes - 1], _cat(m, comum).vendas_limite, _cat(m, comum).isento, m.isento_pme,
         m.perda_pme)
        for m in meses
    ]  # fmt: skip
    outras = Secao(
        "Outras informações (só para consulta)",
        colunas,
        tuple(linhas),
        _total(linhas, colunas, {1, 2, 3, 4}) if linhas else None,
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
        (*mercados, *resultados, consolidacao, outras, fii),
        pdf=True,
        notas=(
            "Mesma ordem dos blocos do programa da Receita; as fichas do ano ficam na tela "
            "Declaração anual.",
        ),
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
        e
        for e in d.apuracao.exclusiva
        if e.titular == d.titular and e.data.year == d.ano and e.tipo is TipoAtivo.ETF_RF
    ]
    infra = [
        e
        for e in d.apuracao.exclusiva
        if e.titular == d.titular and e.data.year == d.ano and e.tipo is TipoAtivo.FUNDO_INFRA
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
    linhas_infra: list[tuple[Valor, ...]] = [
        (e.data, e.ticker, e.quantidade, e.alienacao, e.custo, e.resultado) for e in infra
    ]
    aliquota_zero = Secao(
        "Alíquota zero (FI-Infra e FIP-IE)",
        colunas,
        tuple(linhas_infra),
        _total(linhas_infra, colunas, {3, 4, 5}) if linhas_infra else None,
        nota="Lei 12.431/2011 art. 3º e Lei 11.478/2007 art. 2º: ganho de pessoa física com "
        "alíquota zero; fora da apuração mensal, do limite de R$ 20 mil e do DARF. O prejuízo "
        "não compensa outros ganhos (IN RFB 1.585/2015 art. 64).",
    )
    return Relatorio(
        "irrf-e-exclusiva",
        "IRRF do ano e tributação exclusiva",
        _cab(d),
        (irrf, exclusiva, *((aliquota_zero,) if linhas_infra else ())),
        pdf=True,
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
        nota="Perda em ação de PME não entra: não compensa nada (IN RFB 1.585/2015 art. 67 §2º); "
        "fica em Outras informações da ficha mensal do IRPF.",
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
