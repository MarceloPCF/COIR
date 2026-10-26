"""Texto da apuração para a CLI: o mês com a memória de cálculo (regra 12) e o resumo do ano."""

from datetime import date
from decimal import Decimal

from coir.domain.fiscal import (
    Apuracao,
    Categoria,
    Grupo,
    MesApurado,
    Movimento,
    PosicaoAtivo,
    Visao,
)

NOMES = {Categoria.COMUM: "Comum", Categoria.DAY_TRADE: "Day trade", Categoria.FII: "FII"}


def _r(valor: Decimal) -> str:
    texto = f"{valor:,.2f}"
    return texto.replace(",", "_").replace(".", ",").replace("_", ".")


def _meses(ap: Apuracao, ano: int, mes: int) -> list[MesApurado]:
    return [m for m in ap.meses if (m.ano, m.mes) == (ano, mes)]


def resumo_mes(ap: Apuracao, ano: int, mes: int, codigo_darf: str) -> str:
    meses = _meses(ap, ano, mes)
    if not meses:
        return f"{mes:02d}/{ano}: nada apurado"
    linhas: list[str] = []
    for m in meses:
        linhas.append(f"== {mes:02d}/{ano} - titular {m.titular}")
        for c in m.categorias:
            if not (c.resultado or c.prejuizo_anterior or c.isento):
                continue
            grupos = ", ".join(f"{g.value} {_r(v)}" for g, v in c.por_grupo.items() if v)
            linhas.append(f"{NOMES[c.categoria]}: resultado {_r(c.resultado)} ({grupos or '-'})")
            if c.categoria is Categoria.COMUM:
                linhas.append(
                    f"  vendas de ações à vista {_r(c.vendas_acoes)}; isento {_r(c.isento)}"
                )
            linhas.append(
                f"  prejuízo anterior {_r(c.prejuizo_anterior)}; base {_r(c.base)} x "
                f"{(c.aliquota * 100).normalize():f}% = imposto {_r(c.imposto)}; "
                f"prejuízo a compensar {_r(c.prejuizo_a_compensar)}"
            )
        linhas.append(
            f"Imposto {_r(m.imposto)} - IRRF {_r(m.irrf_deduzido)} (do mês {_r(m.irrf_mes)}, "
            f"saldo anterior {_r(m.irrf_saldo_anterior)}; saldo {_r(m.irrf_saldo)}) = "
            f"{_r(m.imposto_devido)}"
        )
        vencimento = m.vencimento.strftime("%d/%m/%Y") if m.vencimento else "-"
        linhas.append(
            f"DARF {codigo_darf}: {_r(m.darf)} (vencimento {vencimento}; acumulado anterior "
            f"{_r(m.darf_acumulado_anterior)}; passa ao mês seguinte {_r(m.darf_acumulado)})"
        )
        if m.isento_pme:
            linhas.append(f"Isento (ações de PME, Lei 13.043/2014): {_r(m.isento_pme)}")
        linhas.append("Memória de cálculo (movimentos com resultado):")
        movs = [x for x in ap.movimentos if x.titular == m.titular and (x.data.year, x.data.month)
                == (ano, mes) and (x.resultado or x.venda_acao_a_vista)]  # fmt: skip
        linhas += [_movimento(x) for x in movs]
    return "\n".join(linhas)


def _movimento(m: Movimento) -> str:
    grupo = "" if m.grupo is Grupo.A_VISTA else f" [{m.grupo.value}]"
    return (
        f"  {m.data:%d/%m} {m.corretora or 'várias':<8} {m.ticker:<10} "
        f"{NOMES[m.categoria]:<9}{grupo} q {m.quantidade:g} "
        f"valor {_r(m.valor)} custo {_r(m.custo_baixa)} resultado {_r(m.resultado)} "
        f"({m.regra}; {m.origem})"
    )


def carteira_texto(
    posicoes: list[PosicaoAtivo], contas: dict[int, str], dia: date, visao: Visao
) -> str:
    if visao is Visao.TITULAR:
        titulo = "PM fiscal por titular (todas as corretoras e contas): apuração e IRPF"
    else:
        titulo = f"Visão operacional por {visao.value}: não é o PM do imposto"
    linhas = [f"Carteira em {dia:%d/%m/%Y} - {titulo}"]
    for p in posicoes:
        onde = p.corretora or ""
        if p.conta is not None:
            onde += f" {contas.get(p.conta, p.conta)}"
        venc = f" venc. {p.vencimento:%d/%m/%Y}" if p.vencimento else ""
        linhas.append(
            f"  titular {p.titular} {onde:<16} {p.ticker:<10}{venc} q {p.quantidade:g} "
            f"PM {p.pm:.8f} custo {_r(p.custo)}"
        )
    if not posicoes:
        linhas.append("  (sem posições abertas)")
    return "\n".join(linhas)


def resumo_ano(ap: Apuracao, ano: int) -> str:
    linhas = [f"Ano {ano}: mês | comum | day trade | FII | imposto | IRRF | DARF | vencimento"]
    for m in (x for x in ap.meses if x.ano == ano):
        res = {c.categoria: c.resultado for c in m.categorias}
        venc = m.vencimento.strftime("%d/%m/%Y") if m.vencimento else "-"
        linhas.append(
            f"{m.mes:02d} | {_r(res[Categoria.COMUM])} | {_r(res[Categoria.DAY_TRADE])} | "
            f"{_r(res[Categoria.FII])} | {_r(m.imposto)} | {_r(m.irrf_deduzido)} | {_r(m.darf)} | "
            f"{venc}"
        )
    for i in (x for x in ap.irrf_anual if x.ano == ano):
        linhas.append(
            f"IRRF não deduzido no ano: 0,005% {_r(i.saldo_comum)} (compensar na declaração); "
            f"day trade {_r(i.saldo_day_trade)} (só restituição). Retido no ano: 0,005% "
            f"{_r(i.comum)}; day trade {_r(i.day_trade)}"
        )
    for e in (x for x in ap.exclusiva if x.data.year == ano):
        linhas.append(f"Tributação exclusiva (ETF de renda fixa): {e.data:%d/%m} {e.ticker} "
                      f"resultado {_r(e.resultado)}")  # fmt: skip
    return "\n".join(linhas)
