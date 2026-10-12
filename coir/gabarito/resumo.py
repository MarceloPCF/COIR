"""Resumo do gabarito para validação humana (sem CPF nem número de conta)."""

from collections import Counter, defaultdict
from decimal import Decimal

from coir.gabarito.modelo import Categoria, Gabarito

DUAS = Decimal("0.01")


def _r(valor: Decimal) -> str:
    texto = f"{valor.quantize(DUAS):,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def gerar(gabarito: Gabarito) -> str:
    ops = gabarito.operacoes
    anos = sorted({o.data.year for o in ops})
    por_ano: dict[int, dict[str, Decimal]] = defaultdict(lambda: defaultdict(Decimal))
    notas: dict[int, set[tuple[str, str, str]]] = defaultdict(set)
    contagem: dict[int, Counter[str]] = defaultdict(Counter)
    for o in ops:
        a = o.data.year
        contagem[a][o.categoria.value] += 1
        notas[a].add((o.corretora, o.conta, o.nota))
        por_ano[a]["compras" if o.cv == "C" else "vendas"] += o.total
        por_ano[a]["custos"] += o.custos
        por_ano[a]["irrf_notas"] += o.irrf
    for m in gabarito.apuracao:
        chave = "resultado_comum" if m.categoria is Categoria.COMUM else "resultado_dt"
        por_ano[m.ano][chave] += m.resultado
        if m.a_pagar > 0:
            por_ano[m.ano]["imposto"] += m.a_pagar
    for d in gabarito.darfs:
        por_ano[d.ano]["darfs"] += d.valor

    linhas = [
        "# Resumo do gabarito (para validação)",
        "",
        "Valores em R$, extraídos da planilha-gabarito. Titulares e contas pseudonimizados.",
        "",
        "## Por ano",
        "",
        "| Ano | Ops comuns | Ops DT | Notas | Compras | Vendas | Custos | IRRF (notas) "
        "| Resultado comum | Resultado DT | Imposto (A pagar > 0) | DARFs lançados |",
        "|---|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|",
    ]
    for a in anos:
        v = por_ano[a]
        linhas.append(
            f"| {a} | {contagem[a]['COMUM']} | {contagem[a]['DAY_TRADE']} | {len(notas[a])} "
            f"| {_r(v['compras'])} | {_r(v['vendas'])} | {_r(v['custos'])} | {_r(v['irrf_notas'])} "
            f"| {_r(v['resultado_comum'])} | {_r(v['resultado_dt'])} | {_r(v['imposto'])} "
            f"| {_r(v['darfs'])} |"
        )
    linhas += ["", "## Operações por corretora/conta e mercado", ""]
    por_conta = Counter((o.conta, o.categoria.value, o.mercado) for o in ops)
    linhas += ["| Conta | Categoria | Mercado | Operações |", "|---|---|---|--:|"]
    linhas += [f"| {c} | {cat} | {m} | {n} |" for (c, cat, m), n in sorted(por_conta.items())]
    linhas += ["", "## Eventos lançados (sem nota)", ""]
    linhas += [f"- {t}: {n}" for t, n in sorted(Counter(e.tipo for e in gabarito.eventos).items())]
    linhas += ["", "## Avisos da extração", ""]
    linhas += [f"- {a}" for a in gabarito.avisos] or ["- nenhum"]
    linhas += [
        "",
        "## Pontos para você decidir na validação",
        "",
        "1. O gabarito calcula o preço médio por papel **e conta** (`Papel_Conta`); a regra "
        "fiscal usa o titular (todas as corretoras). Confirmar se é intencional.",
        "2. A aba Isentos está zerada a partir de 2021 — conferir se é real ou erro da planilha.",
        "3. A aba BTC_Outros não tem valores mensais — onde estão os lançamentos de outros custos?",
        "4. A aba Custodia_Inicial não tem posições — confirmar que o histórico começa sem "
        "custódia.",
        "5. Opções lançadas (venda a descoberto) e alguns eventos têm o resultado apurado no mês "
        "de encerramento/vencimento, não no do pregão (207 linhas). O COIR v3 seguirá a mesma "
        "regra se confirmada pelo especialista fiscal.",
        "6. Na apuração comum, as vendas marcadas 'Isento = Sim' (mês com vendas de ações até "
        "R$ 20 mil) ficam fora do resultado — inclusive **prejuízos** (ex.: fracionário em "
        "06/2008). Prejuízo em mês isento pode ser compensado? Ponto a decidir na F4.",
        "7. O gabarito registra IRRF de até R$ 1,00 por nota, mas o PDF mostra que ele não foi "
        "retido (dispensado). Na F4, IRRF dispensado não deve virar crédito a deduzir.",
        "",
    ]
    return "\n".join(linhas)
