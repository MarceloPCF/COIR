"""Relatórios do patrimônio: carteira numa data, bens e direitos e proventos recebidos."""

from collections import defaultdict
from collections.abc import Callable
from dataclasses import replace

from coir.domain.fiscal import (
    PosicaoAtivo,
    TipoAtivo,
    Visao,
)
from coir.domain.proventos import Provento, TipoProvento
from coir.domain.relatorio import DadosRelatorio
from coir.reports.fiscais.comum import T, _cab, _dinheiro, _soma, _total
from coir.reports.modelo import (
    Coluna,
    Relatorio,
    Secao,
    Valor,
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
    ("FI-Infra e FIP-IE", lambda p: p.quantidade > 0 and p.tipo is TipoAtivo.FUNDO_INFRA),
    ("FIP e FIDC", lambda p: p.quantidade > 0 and p.tipo in (TipoAtivo.FIP, TipoAtivo.FIDC)),
    ("Opções", lambda p: p.quantidade > 0 and p.tipo is TipoAtivo.OPCAO),
    (
        "Direitos, recibos e bônus de subscrição",
        lambda p: p.quantidade > 0 and p.tipo is TipoAtivo.DIREITO,
    ),
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


# --------------------------------------------------------------------------------- 12. proventos
NOME_PROVENTO = {
    TipoProvento.DIVIDENDO: "Dividendo", TipoProvento.JCP: "JCP",
    TipoProvento.RENDIMENTO_FII: "Rendimento de FII", TipoProvento.REEMBOLSO_BTC: "Reembolso BTC",
    TipoProvento.RENDIMENTO_FUNDO_INFRA: "Rendimento de FI-Infra/FIP-IE",
    TipoProvento.ATUALIZACAO: "Atualização (Selic)",
    TipoProvento.RENDIMENTO_ETF: "Rendimento de ETF",
    TipoProvento.RENDIMENTO_ETF_RF: "Rendimento de ETF de renda fixa",
    TipoProvento.RESTITUICAO_CAPITAL: "Restituição de capital",
    TipoProvento.FORA_DO_ESCOPO: "Só aviso (o COIR não calcula)",
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
    entra no ano do crédito, critério do COIR de 19/10/2026)."""
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
        (TipoProvento.RENDIMENTO_FUNDO_INFRA,
         "Rendimentos de FI-Infra e FIP-IE (tributação exclusiva, código 06)", False,
         "Lei 12.431/2011 art. 3º §§1º e 2º; Lei 11.478/2007 art. 2º §3º: alíquota zero para "
         "pessoa física, pelo valor líquido. Vale o informe do administrador."),
        (TipoProvento.RENDIMENTO_ETF,
         "Rendimentos distribuídos por ETF (tributação exclusiva, código 06)", True,
         "Lei 14.754/2023 arts. 18 II, 24 e 32 I: IRRF de 15% na distribuição, definitivo para a "
         "pessoa física, desde 2024. Pelo valor líquido; vale o informe do administrador."),
        (TipoProvento.RENDIMENTO_ETF_RF,
         "Rendimentos distribuídos por ETF de renda fixa (tributação exclusiva, código 06)", True,
         "Lei 13.043/2014 art. 2º §4º; IN RFB 1.585/2015 arts. 29 III e 30 II: exclusiva na "
         "fonte, retida pelo administrador. Pelo valor líquido; vale o informe."),
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
