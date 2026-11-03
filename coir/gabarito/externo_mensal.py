"""Relatório mensal de outro sistema de IR (PDF "Operações Comuns e Day-Trade", um por ano):
a ficha de renda variável mês a mês, na ordem do programa IRPF, como gabarito independente da
apuração do motor (resultado por mercado, prejuízo a compensar, imposto e IR fonte).

Layout: uma página por mês ("Janeiro 2020 Mesma ordem e nomenclatura do programa"), linhas
"rótulo R$ x" (uma coluna) ou "rótulo R$ x R$ y" (comuns e day trade); negativo como "-R$ x".
O cabeçalho traz nome e CPF do titular: não é lido. Rótulo desconhecido falha alto, para a
comparação nunca rodar sobre uma leitura pela metade.

O primeiro mês de cada relatório traz o "resultado negativo até o mês anterior" que o outro
sistema recebeu (prejuízo declarado): é o saldo inicial com que o motor é comparado
(`saldo_inicial`), para que diferenças de antes do marco não contaminem os anos seguintes.
"""

import csv
import re
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, replace
from decimal import Decimal
from itertools import pairwise
from pathlib import Path

from pydantic import BaseModel, ConfigDict

from coir.domain.fiscal import Categoria, Ficha, Grupo, MesApurado, SaldoInicial
from coir.erros import GabaritoError
from coir.parsing.texto_pdf import paginas

ZERO = Decimal(0)
MESES = ("Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto",
         "Setembro", "Outubro", "Novembro", "Dezembro")  # fmt: skip
RE_VALOR = r"-?R\$ (?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}|\d+%"
RE_MES = re.compile(rf"^(?P<mes>{'|'.join(MESES)}) (?P<ano>\d{{4}}) Mesma ordem")
RE_LINHA = re.compile(rf"^(?P<rotulo>.+?) (?P<v1>{RE_VALOR})(?: (?P<v2>{RE_VALOR}))?$")

# linhas de mercado: rótulo do relatório -> grupos do motor que somam nela
MERCADOS: dict[str, tuple[Grupo, ...]] = {
    "Mercado à vista - ações": (Grupo.A_VISTA,),
    "Mercado à vista - ouro": (),
    "Mercado à vista - ouro at. fin. fora bolsa": (),
    "Mercado opções - ações": (Grupo.OPCOES,),
    "Mercado opções - ouro": (),
    "Mercado opções - fora de bolsa": (),
    "Mercado opções - outros": (Grupo.OPCOES_OUTROS,),
    "Mercado futuro - dólar dos EUA": (Grupo.DOLAR,),
    "Mercado futuro - índices": (Grupo.INDICE,),
    "Mercado futuro - juros": (Grupo.JUROS,),
    "Mercado futuro - outros": (Grupo.COMMODITY,),
    "Mercado a termo - ações/ouro": (Grupo.TERMO,),
    "Mercado a termo - outros": (),
}
RESULTADO = "Resultado líquido do mês"
NEGATIVO_ANTERIOR = "Resultado negativo até o mês anterior"
BASE = "Base de cálculo do imposto"
PREJUIZO = "Prejuízo a compensar"
ALIQUOTA = "Alíquota do imposto"
IMPOSTO = "Imposto devido"
TOTAL_IMPOSTO = "Total do imposto devido"
IRRF_DT = "IR fonte de Day Trade no mês"
IRRF_DT_ANTERIOR = "IR fonte de Day Trade nos meses anteriores"
IRRF_DT_COMPENSAR = "IR fonte de Day Trade a compensar"
IRRF = "IR fonte (Lei nr 11.033/2004)"
IRRF_ANTERIOR = "IR fonte (Lei nr 11.033/2004) nos meses anteriores"
IRRF_COMPENSAR = "IR fonte (Lei nr 11.033/2004) a compensar"
A_PAGAR = "Imposto a pagar"
VENDAS = "Total em alienações de ações (op. comuns) (Lei nr 11.033/2004)"
ISENTO = "Lucro isento (Lei nr 11.033/2004)"
ISENTO_PME = "Lucro isento (Lei nr 13.043/2014)"
# linhas lidas mas não comparadas (o motor não tem o caso)
SEM_COMPARACAO = (ALIQUOTA, "Lucro isento (Lei nr 11.478/2007)",
                  "Lucro isento em ouro - ativo financeiro")  # fmt: skip
CONHECIDAS = {*MERCADOS, RESULTADO, NEGATIVO_ANTERIOR, BASE, PREJUIZO, IMPOSTO, TOTAL_IMPOSTO,
              IRRF_DT, IRRF_DT_ANTERIOR, IRRF_DT_COMPENSAR, IRRF, IRRF_ANTERIOR, IRRF_COMPENSAR,
              A_PAGAR, VENDAS, ISENTO, ISENTO_PME, *SEM_COMPARACAO}  # fmt: skip


class MesExterno(BaseModel):
    """Um mês do relatório: rótulo -> (comuns, day trade); day trade None nas linhas de uma
    coluna só (consolidação do mês e informações para consulta)."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    ano: int
    mes: int
    valores: dict[str, tuple[Decimal, Decimal | None]]

    def valor(self, rotulo: str, coluna: int = 0) -> Decimal:
        v = self.valores.get(rotulo, (None, None))[coluna]
        if v is None:
            raise GabaritoError(f"relatório externo {self.mes:02d}/{self.ano}: sem {rotulo!r}")
        return v


def _valor(texto: str) -> Decimal:
    if texto.endswith("%"):
        return Decimal(texto[:-1]) / 100
    negativo = texto.startswith("-")
    numero = Decimal(texto.lstrip("-").removeprefix("R$ ").replace(".", "").replace(",", "."))
    return -numero if negativo else numero


def ler_texto(textos: Sequence[str]) -> list[MesExterno]:
    meses: list[MesExterno] = []
    for texto in textos:
        linhas = [x.strip() for x in texto.splitlines() if x.strip()]
        cabecalho = next((m for x in linhas if (m := RE_MES.match(x))), None)
        if cabecalho is None:
            continue  # capa, isentos do ano, aviso
        valores: dict[str, tuple[Decimal, Decimal | None]] = {}
        for x in linhas:
            if (m := RE_LINHA.match(x)) is None:
                continue
            if m["rotulo"] not in CONHECIDAS:
                raise GabaritoError(f"relatório externo: linha desconhecida {m['rotulo']!r}")
            valores[m["rotulo"]] = (_valor(m["v1"]), _valor(m["v2"]) if m["v2"] else None)
        faltam = CONHECIDAS - set(valores)
        if faltam:
            raise GabaritoError(f"relatório externo {cabecalho[0]}: faltam {sorted(faltam)}")
        meses.append(MesExterno(ano=int(cabecalho["ano"]), mes=MESES.index(cabecalho["mes"]) + 1,
                                valores=valores))  # fmt: skip
    if not meses:
        raise GabaritoError("relatório externo sem nenhum mês")
    return meses


def ler_relatorios(arquivos: Iterable[Path]) -> list[MesExterno]:
    """Meses de vários relatórios (um por ano), em ordem; mês repetido ou faltando falha."""
    meses = sorted((m for a in arquivos for m in ler_texto(paginas(a))),
                   key=lambda m: (m.ano, m.mes))  # fmt: skip
    chaves = [(m.ano, m.mes) for m in meses]
    if len(set(chaves)) != len(chaves):
        raise GabaritoError("relatório externo: mês repetido entre os arquivos")
    for (a, n), (b, k) in pairwise(chaves):
        if (b, k) != ((a, n + 1) if n < 12 else (a + 1, 1)):
            raise GabaritoError(f"relatório externo: falta o mês seguinte a {n:02d}/{a}")
    return meses


def saldo_inicial(meses: Sequence[MesExterno], titular: int) -> SaldoInicial:
    """Prejuízo que o relatório recebeu no primeiro mês (31/12 do ano anterior)."""
    primeiro = meses[0]
    if primeiro.mes != 1:
        raise GabaritoError("relatório externo deve começar em janeiro")
    return SaldoInicial(titular=titular, ano=primeiro.ano - 1,
                        comum=primeiro.valor(NEGATIVO_ANTERIOR),
                        day_trade=primeiro.valor(NEGATIVO_ANTERIOR, 1))  # fmt: skip


@dataclass(frozen=True)
class LinhaMensal:
    ano: int
    mes: int
    rotulo: str
    coluna: str  # comuns, day_trade
    motor: Decimal
    externo: Decimal
    classe: str  # confere, divergente, explicada
    motivo: str = ""

    @property
    def diferenca(self) -> Decimal:
        return self.motor - self.externo


def _motor(m: MesApurado) -> dict[str, tuple[Decimal, Decimal | None]]:
    """O mês do motor com os rótulos do relatório (prejuízos positivos, como no programa)."""
    cat = {c.categoria: c for c in m.categorias}
    comum, dt = cat[Categoria.COMUM], cat[Categoria.DAY_TRADE]
    ficha = m.ficha(Ficha.COMUNS)
    saida: dict[str, tuple[Decimal, Decimal | None]] = {
        rotulo: (sum((comum.por_grupo.get(g, ZERO) for g in grupos), ZERO),
                 sum((dt.por_grupo.get(g, ZERO) for g in grupos), ZERO))
        for rotulo, grupos in MERCADOS.items()
    }  # fmt: skip
    saida |= {
        RESULTADO: (comum.resultado, dt.resultado),
        NEGATIVO_ANTERIOR: (-comum.prejuizo_anterior, -dt.prejuizo_anterior),
        BASE: (comum.base, dt.base),
        PREJUIZO: (-comum.prejuizo_a_compensar, -dt.prejuizo_a_compensar),
        IMPOSTO: (comum.imposto, dt.imposto),
        TOTAL_IMPOSTO: (ficha.imposto, None),
        IRRF_DT: (ficha.irrf_day_trade_mes, None),
        IRRF_DT_ANTERIOR: (ficha.irrf_day_trade_anterior, None),
        IRRF_DT_COMPENSAR: (ficha.irrf_day_trade_a_compensar, None),
        IRRF: (ficha.irrf_mes, None),
        IRRF_ANTERIOR: (ficha.irrf_anterior, None),
        IRRF_COMPENSAR: (ficha.irrf_a_compensar, None),
        A_PAGAR: (ficha.a_pagar, None),
        VENDAS: (comum.vendas_acoes, None),
        ISENTO: (comum.isento, None),
        ISENTO_PME: (m.isento_pme, None),
    }
    return saida


def comparar_mensal(meses_externos: Iterable[MesExterno], apurados: Iterable[MesApurado],
                    titular: int, tolerancia: Decimal) -> list[LinhaMensal]:  # fmt: skip
    """Linha a linha (comuns e day trade) dos meses do relatório. O motor apura todos os meses
    da primeira operação do titular em diante: mês do relatório que ele não apurou (titular
    errado, ou relatório que começa antes das notas) falha alto, em vez de comparar com zero."""
    motor = {(m.ano, m.mes): _motor(m) for m in apurados if m.titular == titular}
    saida: list[LinhaMensal] = []
    for ext in meses_externos:
        meu = motor.get((ext.ano, ext.mes))
        if meu is None:
            raise GabaritoError(f"o motor não apurou {ext.mes:02d}/{ext.ano} do titular {titular} "
                                "(sem notas até esse mês?)")  # fmt: skip
        for rotulo, valores in ext.valores.items():
            if rotulo in SEM_COMPARACAO:
                continue
            for coluna, nome in ((0, "comuns"), (1, "day_trade")):
                externo = valores[coluna]
                if externo is None:
                    continue
                v = meu[rotulo][coluna]
                if v is None:
                    raise GabaritoError(f"linha {rotulo!r} sem a coluna {nome} no motor")
                classe = "confere" if abs(v - externo) <= tolerancia else "divergente"
                saida.append(LinhaMensal(ext.ano, ext.mes, rotulo, nome, v, externo, classe))
    return saida


@dataclass(frozen=True)
class Explicacao:
    """Divergência já analisada; "*" em mes, rotulo ou coluna vale para todos (ex.: o prejuízo
    acumulado que carrega uma diferença explicada nos meses seguintes)."""

    ano: int
    mes: str
    rotulo: str
    coluna: str
    motivo: str

    def cobre(self, x: LinhaMensal) -> bool:
        return (self.ano == x.ano and self.mes in ("*", str(x.mes))
                and self.rotulo in ("*", x.rotulo) and self.coluna in ("*", x.coluna))  # fmt: skip


def carregar_explicacoes(arquivo: Path) -> list[Explicacao]:
    """CSV em COIR_DADOS com colunas ano, mes, rotulo, coluna, motivo. Ausente: nenhuma."""
    if not arquivo.is_file():
        return []
    with arquivo.open(encoding="utf-8-sig", newline="") as f:
        linhas = list(csv.DictReader(f))
    try:
        return [Explicacao(int(r["ano"]), r["mes"].strip(), r["rotulo"].strip(),
                           r["coluna"].strip(), r["motivo"].strip()) for r in linhas]  # fmt: skip
    except (KeyError, ValueError, AttributeError) as erro:
        raise GabaritoError(f"explicações inválidas em {arquivo.name}: {erro}") from erro


def explicar(linhas: Iterable[LinhaMensal], explicacoes: Sequence[Explicacao]
             ) -> list[LinhaMensal]:  # fmt: skip
    """Marca como explicada a divergência coberta pela primeira explicação que vale para ela."""
    saida: list[LinhaMensal] = []
    for x in linhas:
        e = next((e for e in explicacoes if e.cobre(x)), None) if x.classe == "divergente" else None
        saida.append(replace(x, classe="explicada", motivo=e.motivo) if e else x)
    return saida


def gravar(linhas: Iterable[LinhaMensal], arquivo: Path) -> None:
    with arquivo.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["ano", "mes", "rotulo", "coluna", "motor", "externo", "diferenca", "classe",
                    "motivo"])  # fmt: skip
        for x in linhas:
            w.writerow([x.ano, x.mes, x.rotulo, x.coluna, x.motor, x.externo, x.diferenca,
                        x.classe, x.motivo])  # fmt: skip
