"""Comparação das fichas do COIR com a declaração entregue (critério de saída da F8).

Compara o que o COIR preenche (`FichasIrpf`) com o que foi declarado (`Dirpf`, lido do PDF):
- renda variável mês a mês, linha a linha (comuns e day trade), com a tolerância de mês;
- Bens e Direitos negociados em bolsa por código de negociação (itens da DIRPF do mesmo ativo em
  corretoras diferentes somados: o COIR gera um item por ativo) e JCP a receber (99.07) por CNPJ;
- rendimentos por linha e CNPJ. Nas linhas que também recebem rendimentos fora da bolsa (99, 06:
  `parcial` no mapeamento) só entram as fontes que o COIR conhece;
- IR fonte da Lei 11.033/2004 de "Imposto pago/retido".

Diferença acima da tolerância é divergência a explicar (`dirpf-explicadas.csv` em COIR_DADOS),
nunca motivo para ajustar o motor (CLAUDE.md, regra 14).
"""

import csv
from collections import defaultdict
from collections.abc import Callable, Iterable
from dataclasses import dataclass, replace
from decimal import Decimal
from pathlib import Path

from coir.domain.irpf import FichasIrpf, MapeamentoIrpf
from coir.gabarito.dirpf import MESES, Dirpf

ZERO = Decimal(0)
Chave = tuple[str, str, str]  # seção, item, coluna


@dataclass(frozen=True)
class LinhaComparacao:
    secao: str  # renda_variavel, bens, rendimentos, imposto_pago
    item: str  # "JAN Mercado à Vista - Ações", "BBAS3", "isentos 09 33592510000154"...
    coluna: str  # comuns, day_trade, anterior, atual, valor, codigo
    coir: Decimal | str | None
    dirpf: Decimal | str | None
    classe: str  # confere, divergente, so_coir, so_dirpf, explicada
    explicacao: str = ""

    @property
    def chave(self) -> Chave:
        return (self.secao, self.item, self.coluna)

    @property
    def diferenca(self) -> Decimal | None:
        if isinstance(self.coir, Decimal) and isinstance(self.dirpf, Decimal):
            return self.coir - self.dirpf
        return None


def _linha(secao: str, item: str, coluna: str, coir: Decimal | str | None,
           dirpf: Decimal | str | None, tolerancia: Decimal) -> LinhaComparacao:  # fmt: skip
    if coir is None:
        classe = "so_dirpf"
    elif dirpf is None:
        classe = "so_coir"
    elif isinstance(coir, Decimal) and isinstance(dirpf, Decimal):
        classe = "confere" if abs(coir - dirpf) <= tolerancia else "divergente"
    else:
        classe = "confere" if coir == dirpf else "divergente"
    return LinhaComparacao(secao, item, coluna, coir, dirpf, classe)


def _renda_variavel(f: FichasIrpf, d: Dirpf, tol: Decimal) -> list[LinhaComparacao]:
    saida: list[LinhaComparacao] = []
    declarados = {m.mes: m for m in d.renda_variavel}
    for m in f.renda_variavel:
        dm = declarados.get(m.mes)
        for x in m.linhas:
            valores = dm.valores.get(x.rotulo) if dm is not None else None
            item = f"{MESES[m.mes - 1]} {x.rotulo}"
            saida.append(_linha("renda_variavel", item, "comuns", x.comuns,
                                valores[0] if valores else None, tol))  # fmt: skip
            if x.day_trade is not None:
                saida.append(_linha("renda_variavel", item, "day_trade", x.day_trade,
                                    valores[1] if valores else None, tol))  # fmt: skip
    return saida


def _bens(f: FichasIrpf, d: Dirpf, tol: Decimal, codigo_atual: Callable[[str], str]
          ) -> list[LinhaComparacao]:  # fmt: skip
    declarados: dict[str, list[Decimal]] = defaultdict(lambda: [ZERO, ZERO])
    codigos: dict[str, set[str]] = defaultdict(set)
    for b in d.bens:
        if b.ticker:
            chave = codigo_atual(b.ticker)  # a DIRPF pode trazer o código antigo (EMBR3 -> EMBJ3)
        elif (b.grupo, b.codigo) == ("99", "07") and b.cnpj:
            chave = f"99.07 {b.cnpj}"
        else:
            continue  # conta bancária, imóvel, renda fixa: fora do COIR
        declarados[chave][0] += b.anterior
        declarados[chave][1] += b.atual
        codigos[chave].add(f"{b.grupo}.{b.codigo}")
    gerados: dict[str, tuple[Decimal, Decimal, str]] = {}
    for b in f.bens:
        chave = codigo_atual(b.ticker) if b.ticker else f"99.07 {b.cnpj}"
        anterior, atual, _ = gerados.get(chave, (ZERO, ZERO, ""))
        gerados[chave] = (anterior + b.anterior, atual + b.atual, f"{b.grupo}.{b.codigo}")
    saida: list[LinhaComparacao] = []
    for chave in sorted(gerados.keys() | declarados.keys()):
        g, dd = gerados.get(chave), declarados.get(chave)
        for i, coluna in enumerate(("anterior", "atual")):
            saida.append(_linha("bens", chave, coluna, g[i] if g else None,
                                dd[i] if dd else None, tol))  # fmt: skip
        if g and dd and g[2] not in codigos[chave]:
            saida.append(_linha("bens", chave, "codigo", g[2], "/".join(sorted(codigos[chave])),
                                tol))  # fmt: skip
    return saida


def _rendimentos(f: FichasIrpf, d: Dirpf, mapa: MapeamentoIrpf, tol: Decimal
                 ) -> list[LinhaComparacao]:  # fmt: skip
    parciais = {("isentos", x.codigo) for x in mapa.isentos.linhas if x.parcial}
    parciais |= {("exclusivos", x.codigo) for x in mapa.exclusivos.linhas if x.parcial}
    comparadas = {("isentos", x.codigo) for x in mapa.isentos.linhas}
    comparadas |= {("exclusivos", x.codigo) for x in mapa.exclusivos.linhas}
    gerados: dict[tuple[str, str, str], Decimal] = defaultdict(lambda: ZERO)
    for r in f.rendimentos:
        if r.ficha != "imposto_pago":
            gerados[(r.ficha, r.codigo, r.cnpj or "")] += r.valor
    declarados: dict[tuple[str, str, str], Decimal] = defaultdict(lambda: ZERO)
    for r in d.rendimentos:
        declarados[(r.ficha, r.codigo, r.cnpj)] += r.valor
    for ficha, codigo in d.totais:  # linha sem fonte (20) ou sem leitura por fonte
        if (ficha, codigo) in comparadas and (ficha, codigo) not in parciais and not any(
                k[:2] == (ficha, codigo) for k in declarados):  # fmt: skip
            declarados[(ficha, codigo, "")] = d.totais[(ficha, codigo)]
    saida: list[LinhaComparacao] = []
    for chave in sorted(gerados.keys() | declarados.keys()):
        ficha, codigo, cnpj = chave
        if (ficha, codigo) not in comparadas:
            continue
        if (ficha, codigo) in parciais and chave not in gerados:
            continue  # fonte fora da bolsa no mesmo código (aplicação financeira, outros)
        item = f"{ficha} {codigo} {cnpj}".strip()
        saida.append(_linha("rendimentos", item, "valor", gerados.get(chave),
                            declarados.get(chave), tol))  # fmt: skip
    for ficha, codigo in sorted(set(d.sem_fonte) & comparadas - parciais):
        total = sum((v for k, v in gerados.items() if k[:2] == (ficha, codigo)), ZERO)
        saida.append(_linha("rendimentos", f"{ficha} {codigo} total", "valor", total,
                            d.totais.get((ficha, codigo)), tol))  # fmt: skip
    return saida


def comparar(f: FichasIrpf, d: Dirpf, mapa: MapeamentoIrpf, tol_mes: Decimal,
             tol_valor: Decimal, codigo_atual: Callable[[str], str] = lambda t: t,
             ) -> list[LinhaComparacao]:  # fmt: skip
    """`tol_mes`: renda variável mês a mês; `tol_valor`: bens, rendimentos e imposto pago;
    `codigo_atual`: código atual de um ticker (tabela de trocas)."""
    if (f.exercicio, f.ano) != (d.exercicio, d.ano):
        raise ValueError(f"fichas de {f.ano} contra a DIRPF do ano-calendário {d.ano}")
    pago = sum((r.valor for r in f.rendimentos if r.ficha == "imposto_pago"), ZERO)
    return [
        *_renda_variavel(f, d, tol_mes), *_bens(f, d, tol_valor, codigo_atual),
        *_rendimentos(f, d, mapa, tol_valor),
        _linha("imposto_pago", "03 Lei 11.033/2004", "valor", pago, d.ir_fonte_lei_11033,
               tol_valor),
    ]  # fmt: skip


COLUNAS_EXPLICACOES = ("secao", "item", "coluna", "motivo")


def carregar_explicacoes(arquivo: Path) -> dict[Chave, str]:
    """Divergências já analisadas (CSV em COIR_DADOS: seção, item, coluna, motivo)."""
    if not arquivo.is_file():
        return {}
    with arquivo.open(encoding="utf-8-sig", newline="") as f:
        return {(r["secao"].strip(), r["item"].strip(), r["coluna"].strip()): r["motivo"]
                for r in csv.DictReader(f)}  # fmt: skip


def explicar(linhas: Iterable[LinhaComparacao], explicacoes: dict[Chave, str]
             ) -> list[LinhaComparacao]:  # fmt: skip
    return [replace(x, classe="explicada", explicacao=explicacoes[x.chave])
            if x.classe != "confere" and x.chave in explicacoes else x for x in linhas]  # fmt: skip


def gravar(linhas: Iterable[LinhaComparacao], arquivo: Path) -> None:
    with arquivo.open("w", encoding="utf-8", newline="") as f:
        escrita = csv.writer(f, delimiter=";", lineterminator="\n")
        escrita.writerow(("secao", "item", "coluna", "coir", "dirpf", "diferenca", "classe",
                          "explicacao"))  # fmt: skip
        for x in linhas:
            escrita.writerow((x.secao, x.item, x.coluna, x.coir if x.coir is not None else "",
                              x.dirpf if x.dirpf is not None else "",
                              x.diferenca if x.diferenca is not None else "", x.classe,
                              x.explicacao))  # fmt: skip
