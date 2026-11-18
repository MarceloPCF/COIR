"""Data da liquidação do termo: vencimento do contrato ou liquidação antecipada pela B3.

A nota traz só o prazo do contrato; a liquidação no vencimento é o pregão + prazo (pregão seguinte
se o dia não for útil). O comprador pode antecipar a liquidação, total ou em parte, e a alienação
do vendedor se dá nessa data (IN RFB 1.585/2015 art. 59 §2º IV: "no vencimento ou liquidação
antecipada"). A movimentação da Área do Investidor B3 registra a entrega das ações numa linha
"Liquidação Termo" (crédito para o comprador, débito para o vendedor), com o código da ação
seguido de "T" no produto. Esta é a única fonte da data antecipada; sem a movimentação do período,
fica o vencimento, com aviso.
"""

import re
import unicodedata
from collections import defaultdict
from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

ZERO = Decimal(0)
RE_TICKER_TERMO = re.compile(r"^([A-Z0-9]{4}\d{1,2})T$")  # PETR4T -> PETR4
LIQUIDACAO_TERMO = "liquidacao termo"


def _normal(texto: str) -> str:
    sem_acento = unicodedata.normalize("NFKD", texto)
    return " ".join("".join(c for c in sem_acento if not unicodedata.combining(c)).lower().split())


def e_liquidacao_termo(movimentacao: str) -> bool:
    return _normal(movimentacao) == LIQUIDACAO_TERMO


def codigos_do_produto(produto: str, codigo: str) -> list[str]:
    """Códigos candidatos da ação numa linha de termo da B3, na ordem de preferência: o código com
    "T" no fim, sem o "T" ("Xxxxx Xxxxx - PETR4T - PETROBRAS", "Termo de Ação PETR4T"), e por
    último o código que o importador gravou."""
    saida: list[str] = []
    for parte in produto.replace(" - ", " ").split():
        m = RE_TICKER_TERMO.match(parte.strip())
        if m and m[1] not in saida:
            saida.append(m[1])
    if codigo and codigo not in saida:
        saida.append(codigo)
    return saida


@dataclass(frozen=True)
class LinhaLiquidacao:
    """Linha "Liquidação Termo" da movimentação da B3, já ligada ao ativo e à corretora."""

    titular: int
    corretora: str | None  # pela tabela de instituições; None se a instituição não tem cadastro
    ativo: int | None  # None se nenhum código do produto é ativo conhecido na data
    entrada: bool  # crédito: o comprador recebe as ações
    data: date
    quantidade: Decimal


@dataclass(frozen=True)
class Contrato:
    """Compra ou venda a termo de uma nota."""

    chave: int  # identificador do contrato (id da operação)
    titular: int
    corretora: str
    ativo: int
    ticker: str
    cv: str
    pregao: date
    vencimento: date
    quantidade: Decimal


@dataclass(frozen=True)
class Liquidacao:
    """Como o contrato liquida: partes (data, quantidade), na ordem, e o aviso, se houver."""

    partes: tuple[tuple[date, Decimal], ...]
    aviso: str = ""


@dataclass
class _Saldo:
    linha: LinhaLiquidacao
    restante: Decimal = field(default=ZERO)


def ler_linhas(
    linhas: Iterable[tuple[int, str, str, str, str, bool, date, Decimal]],
    corretora_b3: Callable[[str], str | None],
    ativo: Callable[[str, date], int | None],
) -> list[LinhaLiquidacao]:
    """(titular, movimentação, produto, código, instituição, entrada, data, quantidade) da
    movimentação da B3 -> linhas de liquidação de termo, ligadas à corretora e ao ativo (None
    quando a instituição ou o código não têm cadastro)."""
    saida: list[LinhaLiquidacao] = []
    for titular, mov, produto, codigo, instituicao, entrada, dia, quantidade in linhas:
        if not e_liquidacao_termo(mov):
            continue
        achado = next((a for c in codigos_do_produto(produto, codigo)
                       if (a := ativo(c, dia)) is not None), None)  # fmt: skip
        saida.append(LinhaLiquidacao(titular=titular, corretora=corretora_b3(instituicao),
                                     ativo=achado, entrada=entrada, data=dia,
                                     quantidade=abs(quantidade)))  # fmt: skip
    return saida


def liquidacoes(
    contratos: Sequence[Contrato],
    linhas: Sequence[LinhaLiquidacao],
    periodo: dict[int, tuple[date, date]],
    vendas: dict[tuple[int, int], list[date]] | None = None,
) -> dict[int, Liquidacao]:
    """Liga cada contrato às linhas "Liquidação Termo" do mesmo titular, corretora, ativo e lado
    (crédito na compra, débito na venda) entre o dia seguinte ao pregão e o vencimento. Os
    contratos do mesmo grupo são atendidos pela ordem do vencimento e do pregão; em cada um, a
    linha com a quantidade exata do contrato vem antes, e as demais na ordem da data. O que a B3
    não mostra liquida no vencimento: com a movimentação do período importada, avisa que a
    liquidação não apareceu nela. Sem ela (parecer do especialista de 17/11/2026), a venda a
    termo sempre avisa (quem antecipa é o comprador); a compra a termo só avisa quando o titular
    vendeu o ativo à vista entre o pregão e o vencimento (``vendas``: datas das vendas por titular
    e ativo) ou quando o vencimento cai em outro ano. Se a venda deixou a posição vendida, o motor
    troca esse aviso por um mais forte (`coir/engine/movimentos.py`), mantendo o vencimento
    (decisão do usuário de 18/11/2026, seguindo o especialista). ``periodo``: primeira e última
    data da
    movimentação importada de cada titular (um intervalo só: um buraco entre duas planilhas não é
    visto, por isso o aviso diz "não encontrou", e não que a B3 não mostra).

    Casamento por heurística quando há dois contratos iguais do mesmo ativo e corretora sem linha
    de quantidade exata: o de vencimento anterior consome primeiro. Linhas a mais (ajuste do
    contrato por provento, sem norma nem caso real) ficam sem uso."""
    grupos: dict[tuple[int, str | None, int | None, bool], list[_Saldo]] = defaultdict(list)
    for linha in sorted(linhas, key=lambda x: x.data):
        grupos[(linha.titular, linha.corretora, linha.ativo, linha.entrada)].append(
            _Saldo(linha, linha.quantidade)
        )
    soltas: dict[int, int] = defaultdict(int)  # linhas sem corretora ou ativo cadastrado
    for linha in linhas:
        if linha.corretora is None or linha.ativo is None:
            soltas[linha.titular] += 1
    saida: dict[int, Liquidacao] = {}
    for c in sorted(contratos, key=lambda c: (c.vencimento, c.pregao, c.chave)):
        saldos = [s for s in grupos.get((c.titular, c.corretora, c.ativo, c.cv == "C"), [])
                  if c.pregao < s.linha.data <= c.vencimento and s.restante > 0]  # fmt: skip
        exata = next((s for s in saldos if s.restante == c.quantidade), None)
        ordem = [exata] if exata else saldos
        falta = c.quantidade
        partes: dict[date, Decimal] = defaultdict(lambda: ZERO)
        for s in ordem:
            if falta <= 0:
                break
            usado = min(falta, s.restante)
            s.restante -= usado
            falta -= usado
            partes[s.linha.data] += usado
        aviso = ""
        if falta > 0:
            partes[c.vencimento] += falta
            vendidas = (vendas or {}).get((c.titular, c.ativo), [])
            aviso = _aviso(c, falta, periodo.get(c.titular), vendidas, soltas.get(c.titular, 0))
        saida[c.chave] = Liquidacao(tuple(sorted(partes.items())), aviso)
    return saida


def _aviso(
    c: Contrato,
    falta: Decimal,
    periodo: tuple[date, date] | None,
    vendas: Sequence[date],
    soltas: int = 0,
) -> str:
    lado = "compra" if c.cv == "C" else "venda"
    qtd = f"{falta.normalize():f}"
    if periodo is not None and periodo[0] <= c.pregao and c.vencimento <= periodo[1]:
        sem_cadastro = (
            f' Há {soltas} linha(s) "Liquidação Termo" que o COIR não ligou a uma corretora '
            "ou a um ativo cadastrado: confira a instituição e o código delas"
            if soltas
            else ""
        )
        return (
            f"Termo: o COIR não encontrou na movimentação da B3 importada a liquidação de {qtd} "
            f"ações da {lado} a termo de {c.pregao:%d/%m/%Y} ({c.ticker}, {c.corretora}) e "
            f"usou o vencimento, {c.vencimento:%d/%m/%Y}. Confira o extrato da corretora e se "
            f"a movimentação importada cobre todo o período do contrato.{sem_cadastro}"
        )
    if c.cv == "C":
        vendeu = sorted(d for d in vendas if c.pregao <= d < c.vencimento)
        if vendeu:
            return (
                f"Termo: houve venda à vista de {c.ticker} em {vendeu[0]:%d/%m/%Y}, antes do "
                f"vencimento da compra a termo de {c.pregao:%d/%m/%Y} ({c.corretora}), e o COIR "
                f"usou o vencimento, {c.vencimento:%d/%m/%Y}, como data de liquidação. Se o "
                "contrato foi liquidado antes, o preço médio dessa venda muda (IN RFB 1.585/2015 "
                "art. 62 I e §1º): importe a movimentação da B3 que cobre o período do contrato "
                "para o COIR ler a liquidação antecipada"
            )
        if c.vencimento.year == c.pregao.year:
            return ""
    return (
        f"Termo: o COIR usou o vencimento, {c.vencimento:%d/%m/%Y}, como data de liquidação de "
        f"{qtd} ações da {lado} a termo de {c.pregao:%d/%m/%Y} ({c.ticker}, "
        f"{c.corretora}). Se o contrato foi liquidado antes, o ganho e o preço médio mudam de "
        "data: importe a movimentação da B3 que cobre o período do contrato para o COIR ler a "
        "liquidação antecipada"
    )
