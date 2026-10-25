"""Notas de corretagem fictícias a partir de um roteiro de operações (YAML em tools/roteiros/).

Cada nota sai no layout SINACOR da XP (Bovespa 2024+ e BM&F 2024+), com resumo, custos, IRRF e
líquido calculados para fechar na conciliação do extrator real, e vira PDF (texto em fonte de
largura fixa). Os proventos do roteiro saem numa planilha de movimentação no formato da B3.
Titular, conta e preços são inventados; nenhum dado real.

Uso: uv run python -m tools.notas_ficticias <roteiro.yaml> <pasta>
"""

import csv
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import date, timedelta
from decimal import ROUND_DOWN, ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Any

import yaml

from coir import referencias
from coir.engine.calendario import dia_de_pregao, vencimento_opcao
from coir.recursos import pasta_data
from tests.apoio import cpf_de_teste

ROTEIRO_DEMONSTRACAO = Path(__file__).resolve().parent / "roteiros" / "demonstracao.yaml"
CENTAVO = Decimal("0.01")
MERCADOS = {"VISTA": "VISTA", "CALL": "OPCAO DE COMPRA", "PUT": "OPCAO DE VENDA",
            "EXERC_CALL": "EXERC OPC COMPRA", "EXERC_PUT": "EXERC OPC VENDA"}  # fmt: skip
OPCOES = ("CALL", "PUT")
INSTITUICAO_B3 = "XP INVESTIMENTOS CCTVM S/A"


# --------------------------------------------------------------------------------- roteiro
@dataclass(frozen=True)
class Negocio:
    cv: str
    mercado: str  # chave de MERCADOS (Bovespa) ou "FUTURO" (BM&F)
    papel: str  # ticker, série da opção ("PETRC380 PN") ou contrato ("WINM25")
    quantidade: int
    preco: Decimal
    day_trade: bool

    @property
    def valor(self) -> Decimal:
        return (self.quantidade * self.preco).quantize(CENTAVO, ROUND_HALF_UP)


@dataclass(frozen=True)
class Nota:
    data: date
    segmento: str  # "bovespa" | "bmf"
    negocios: tuple[Negocio, ...]
    numero: int


@dataclass(frozen=True)
class Roteiro:
    titular: str
    corretora: str
    conta: str
    custos: dict[str, Decimal]
    irrf: dict[str, Decimal]
    vencimentos_futuros: dict[str, date]
    notas: tuple[Nota, ...]
    proventos: tuple[tuple[date, str, str, int, Decimal, Decimal], ...]  # ..., unitário, creditado
    darfs_pagas: dict[tuple[int, int], date]

    @property
    def cpf(self) -> str:
        return cpf_de_teste()


def _decimal(texto: object) -> Decimal:
    """'137.500,00' -> 137500.00 (formato da nota); '0.00025' fica como está."""
    bruto = str(texto)
    return Decimal(bruto.replace(".", "").replace(",", ".") if "," in bruto else bruto)


def _negocio(linha: str, segmento: str) -> Negocio:
    partes = linha.split()
    day_trade = partes[-1] == "D"
    if day_trade:
        partes.pop()
    cv, *meio, quantidade, preco = partes
    if cv not in ("C", "V") or not meio:
        raise ValueError(f"negócio fora do formato: {linha!r}")
    if segmento == "bmf":
        return Negocio(cv, "FUTURO", " ".join(meio), int(quantidade), _decimal(preco), True)
    mercado = meio.pop(0) if meio[0] in MERCADOS else "VISTA"
    return Negocio(cv, mercado, " ".join(meio), int(quantidade), _decimal(preco), day_trade)


def carregar(caminho: Path = ROTEIRO_DEMONSTRACAO) -> Roteiro:
    bruto: dict[str, Any] = yaml.safe_load(caminho.read_text(encoding="utf-8"))
    cal = referencias.carregar_fiscais().calendario
    notas: list[Nota] = []
    for i, n in enumerate(bruto["notas"]):
        dia: date = n["data"]
        if not dia_de_pregao(dia, cal):
            raise ValueError(f"{dia:%d/%m/%Y} não é dia de pregão")
        segmento = n.get("segmento", "bovespa")
        negocios = tuple(_negocio(linha, segmento) for linha in n["negocios"])
        notas.append(Nota(dia, segmento, negocios, int(bruto["primeira_nota"]) + i))
    return Roteiro(
        titular=bruto["titular"],
        corretora=bruto["corretora"],
        conta=str(bruto["conta"]),
        custos={k: Decimal(str(v)) for k, v in bruto["custos"].items()},
        irrf={k: Decimal(str(v)) for k, v in bruto["irrf"].items()},
        vencimentos_futuros=dict(bruto.get("vencimentos_futuros", {})),
        notas=tuple(notas),
        proventos=tuple(
            (
                p[0],
                p[1],
                p[2],
                int(p[3]),
                _decimal(p[4]),
                _decimal(p[5]) if len(p) > 5 else int(p[3]) * _decimal(p[4]),
            )
            for p in bruto.get("proventos", [])
        ),
        darfs_pagas={(int(k[:4]), int(k[5:])): v for k, v in bruto.get("darfs_pagas", {}).items()},
    )


# --------------------------------------------------------------------------------- formatos
def brl(valor: Decimal) -> str:
    """Decimal('1234.5') -> '1.234,50' (sem sinal)."""
    inteiro, centavos = f"{abs(valor).quantize(CENTAVO, ROUND_HALF_UP):.2f}".split(".")
    return f"{int(inteiro):,}".replace(",", ".") + "," + centavos


def _qtd(quantidade: int) -> str:
    return f"{quantidade:,}".replace(",", ".")


def _dc(valor: Decimal) -> str:
    return "C" if valor >= 0 else "D"


def _taxa(volume: Decimal, aliquota: Decimal) -> Decimal:
    return (volume * aliquota).quantize(CENTAVO, ROUND_HALF_UP)


def _truncado(valor: Decimal) -> Decimal:
    return valor.quantize(CENTAVO, ROUND_DOWN)


def _cpf(cpf: str, separador: str = "-") -> str:
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}{separador}{cpf[9:]}"


def _nomes_de_pregao() -> dict[str, str]:
    """Ticker -> nome de pregão (o primeiro da lista pública), para a linha da nota."""
    nomes: dict[str, str] = {}
    with (pasta_data() / "ativos_b3.csv").open(encoding="utf-8") as arquivo:
        for linha in csv.DictReader(arquivo):
            nomes.setdefault(linha["ticker"], linha["nome_pregao"])
    return nomes


def _pregoes_depois(dia: date, quantos: int) -> date:
    cal = referencias.carregar_fiscais().calendario
    while quantos:
        dia += timedelta(days=1)
        if dia_de_pregao(dia, cal):
            quantos -= 1
    return dia


def _prazo_opcao(serie: str, pregao: date) -> str:
    """Mês/ano do vencimento pela letra da série (A-L compra, M-X venda)."""
    mes = (ord(serie[4]) - ord("A")) % 12 + 1
    ano = pregao.year + (1 if mes < pregao.month else 0)
    vencimento = vencimento_opcao(ano, mes, referencias.carregar_fiscais().calendario)
    return f"{vencimento:%m/%y}"


# ------------------------------------------------------------------------------ cabeçalhos
XP_BOVESPA = """XP INVESTIMENTOS CORRETORA DE CÂMBIO, TÍTULOS E VALORES MOBILIÁRIOS S.A.
AV ATAULFO DE PAIVA, 153 - SALA 201 LEBLON 22440-032 RIO DE JANEIRO - RJ
Tel. +55 11 4003-3710 Fax:
Internet: https://www.xpi.com.br/ SAC: 0800-772-0202 e-mail: C.N.P.J: 02.332.886/0001-04 \
Carta Patente: Ouvidoria: Tel. 0800 722 3730 e-mail ouvidoria:
Cliente
Participante destino do repasse
-
Cliente
0
Valor
Banco Agência Conta corrente Acionista Administrador
C.P.F./C.N.P.J/C.V.M./C.O.B.
Código cliente Assessor
0000
Custodiante C.I N
Complemento nome P. Vinc N"""
LEGENDA_BOVESPA = """Especificações diversas
A coluna Q indica liquidação no Agente do Qualificado.
(*) Observações A - Posição futuro T - Liquidação pelo Bruto
2 - Corretora ou pessoa vinculada atuou na
contra parte. C - Clubes e fundos de Ações I - POP
# - Negócio direto P - Carteira Própria
8 - Liquidação Institucional H - Home Broker
D - Day Trade X - Box
F - Cobertura Y - Desmanche de Box
B - Debêntures L - Precatório"""
RODAPE = ("Capitais e regiões metropolitanas: +55 11 4003-3710 Demais localidades: "
          "0800-880-3710 SAC: 0800-772-0202 Ouvidoria: 0800 722 3730")  # fmt: skip
XP_BMF = """XP INVESTIMENTOS CORRETORA DE CÂMBIO, TÍTULOS E VALORES
MOBILIÁRIOS S.A.
Fone: +55 11 4003-3710 Fax:
AV ATAULFO DE PAIVA, 153 - SALA 201 LEBLON 22440-032 RIO DE JANEIRO - RJ
Internet: https://www.xpi.com.br/ SAC: 0800-772-0202 e-mail: Ouvidoria: 0800 722 3730 e-mail \
ouvidoria:
Corretora
Cliente
C.N.P.J.
02.332.886/0001-04
Numero da Corretora
3-5
C.N.P.J/C.P.F
Codigo do Cliente"""
LEGENDA_BMF = """+Custos BM&F, conforme Ofício Circular BM&F 079/2007-DG
-Exercícios de opções=EXO
@ Corretora ou Pessoa Vinculada atuou na Contra Parte
* Negocios gerados automaticamente pelo sistema."""


# ---------------------------------------------------------------------------------- Bovespa
def texto_bovespa(nota: Nota, roteiro: Roteiro, nomes: dict[str, str]) -> str:
    c, ir = roteiro.custos, roteiro.irrf
    linhas_neg: list[str] = []
    vista_c = vista_v = opc_c = opc_v = Decimal(0)
    vendas_comuns = Decimal(0)
    resultado_dt = Decimal(0)
    for n in nota.negocios:
        tokens = n.papel.split()
        if n.mercado in OPCOES:
            especificacao = f"{_prazo_opcao(tokens[0], nota.data)} {n.papel}"
        elif n.mercado == "VISTA":
            especificacao = nomes.get(n.papel, n.papel)
        else:
            especificacao = n.papel
        obs = " D" if n.day_trade else ""
        dc = "D" if n.cv == "C" else "C"
        linhas_neg.append(f"1-BOVESPA {n.cv} {MERCADOS[n.mercado]} {especificacao}{obs} "
                          f"{_qtd(n.quantidade)} {brl(n.preco)} {brl(n.valor)} {dc}")  # fmt: skip
        opcao = n.mercado in OPCOES
        if n.cv == "C":
            opc_c, vista_c = (opc_c + n.valor, vista_c) if opcao else (opc_c, vista_c + n.valor)
        else:
            opc_v, vista_v = (opc_v + n.valor, vista_v) if opcao else (opc_v, vista_v + n.valor)
        if n.day_trade:
            resultado_dt += n.valor if n.cv == "V" else -n.valor
        elif n.cv == "V":
            vendas_comuns += n.valor
    volume = vista_c + vista_v + opc_c + opc_v
    liquido_operacoes = vista_v + opc_v - vista_c - opc_c
    liquidacao = _taxa(volume, c["liquidacao"])
    emolumentos = _taxa(volume, c["emolumentos"])
    corretagem = c["corretagem"]
    irrf = _truncado(vendas_comuns * ir["comum"])
    irrf_debitado = irrf > ir["minimo"]
    irrf_dt = _truncado(resultado_dt * ir["day_trade"]) if resultado_dt > 0 else Decimal(0)
    cblc = liquido_operacoes - liquidacao
    despesas = corretagem + (irrf if irrf_debitado else 0)
    liquido = cblc - emolumentos - despesas - irrf_dt
    so_opcoes = not (vista_c or vista_v)
    liquidacao_em = _pregoes_depois(nota.data, 1 if so_opcoes else 2)
    resumo_negocios = (Decimal(0), vista_v, vista_c, opc_c, opc_v, Decimal(0), Decimal(0), volume)
    irrf_texto = f"{brl(irrf)} D" if irrf_debitado else brl(irrf)
    linhas = [
        "Negociações",
        "Negócios realizados Q Negociação C/V Tipo mercado Prazo Especificação do título Obs. (*) "
        "Quantidade Preço / Ajuste Valor Operação / Ajuste D/C",
        *linhas_neg,
        "NOTA DE NEGOCIAÇÃO", "Nr. nota", str(nota.numero), "Folha", "1", "Data pregão",
        f"{nota.data:%d/%m/%Y}",
        XP_BOVESPA,
        *(brl(v) for v in resumo_negocios),
        "Resumo dos Negócios Debêntures",
        "Vendas à vista Compras à vista Opções - compras Opções - vendas Operações à termo",
        "Valor das oper. c/ títulos públ. (v. nom.)",
        "Valor das operações",
        LEGENDA_BOVESPA,
        f"Total CBLC {brl(cblc)} {_dc(cblc)}",
        f"Valor líquido das operações {brl(liquido_operacoes)} {_dc(liquido_operacoes)}",
        f"Taxa de liquidação {brl(liquidacao)} D",
        "Taxa de Registro 0,00",
        f"Total Bovespa / Soma {brl(emolumentos)} D",
        "Taxa de termo/opções 0,00",
        "Taxa A.N.A. 0,00",
        f"Emolumentos {brl(emolumentos)} D",
        "Taxa de Transf. de Ativos 0,00",
        f"Total Custos / Despesas {brl(despesas)} D",
        f"Taxa Operacional {brl(corretagem)} D",
        "Execução 0,00",
        "Taxa de Custódia 0,00",
        "Impostos 0,00",
        f"I.R.R.F. s/ operações, base R${brl(vendas_comuns)} {irrf_texto}",
        *([f"IRRF Day Trade: Base R${brl(resultado_dt)} Projeção R${brl(irrf_dt)}"]
          if irrf_dt else []),
        "Outros 0,00",
        f"Líquido para {liquidacao_em:%d/%m/%Y} {brl(liquido)} {_dc(liquido)}",
        "Resumo Financeiro Clearing", "Bolsa", "Custos Operacionais",
        "Observação: (1) As operações a termo não são computadas no líquido da", "fatura.",
        RODAPE,
        f"{roteiro.conta} {roteiro.titular.upper()}",
        "Rua Exemplo, 0 - Apto 0 - Bairro Tel. (00) 00000-0000",
        "00000-000 Cidade - UF",
        _cpf(roteiro.cpf),
        f" 3-5 {roteiro.conta} ",
    ]  # fmt: skip
    return "\n".join(linhas) + "\n"


# ------------------------------------------------------------------------------------ BM&F
def texto_bmf(nota: Nota, roteiro: Roteiro) -> str:
    c, ir = roteiro.custos, roteiro.irrf
    ajuste = nota.negocios[-1].preco  # preço de ajuste do dia: o do último negócio
    contratos_ref = referencias.carregar().contratos
    linhas_neg: list[str] = []
    ajuste_dt = Decimal(0)
    contratos = 0
    for n in nota.negocios:
        raiz, serie = n.papel[:3], n.papel[3:]
        multiplicador = next(c.multiplicador for c in contratos_ref
                             if c.raiz == raiz and c.vigente(nota.data))  # fmt: skip
        sinal = 1 if n.cv == "C" else -1
        valor = (sinal * (ajuste - n.preco) * n.quantidade * multiplicador).quantize(CENTAVO)
        ajuste_dt += valor
        contratos += n.quantidade
        vencimento = roteiro.vencimentos_futuros[n.papel]
        linhas_neg.append(f"{n.cv} {raiz} {serie} {vencimento:%d/%m/%Y} {n.quantidade} "
                          f"{brl(n.preco)} DAY TRADE {brl(valor)} {_dc(valor)} 0,00")  # fmt: skip
    registro = contratos * c["bmf_registro_por_contrato"]
    emolumentos = contratos * c["bmf_emolumentos_por_contrato"]
    custos = registro + emolumentos
    antes_irrf = ajuste_dt - custos
    irrf_dt = _truncado(antes_irrf * ir["day_trade"]) if antes_irrf > 0 else Decimal(0)
    liquido = antes_irrf - irrf_dt

    def m(valor: Decimal) -> str:
        return f"{brl(valor)} | {_dc(valor)}"

    linhas = [
        "Negociações",
        "C/V Mercadoria Vencimento Quantidade Preço/Ajuste Tipo Negócio Vlr de Operação/Ajuste D/C "
        "Taxa Operacional",
        *linhas_neg,
        "NOTA DE NEGOCIAÇÃO Nr. nota", str(nota.numero), "Folha", "1", "Data pregão",
        f"{nota.data:%d/%m/%Y}",
        XP_BMF,
        roteiro.conta,
        "Venda disponível Compra disponível Venda Opções Compra Opções Valor dos negócios",
        f"0,00 0,00 0,00 0,00 {m(ajuste_dt)}",
        "IRRF IRRF Day Trade (proj.) Taxa operacional Taxa registro BM&F Taxas BM&F (emol+f.gar)",
        f"0,00| {brl(irrf_dt)} 0,00 {brl(registro)} {brl(emolumentos)} | D",
        "+Outros Custos Impostos Ajuste de posição Ajuste day trade Total de custos operacionais",
        f"0,00 0,00 0,00 | C {m(ajuste_dt)} {brl(custos)} | D",
        "Outros IRRF operacional Total Conta Investimento Total Conta Normal Total liquido (#) "
        "Total líquido da nota",
        f"0,00 0,00 0,00| {m(liquido)} {m(max(antes_irrf, Decimal(0)))} {m(liquido)}",
        LEGENDA_BMF,
        RODAPE,
        roteiro.titular.upper(),
        "Rua Exemplo, 0 - Apto 0",
        "00000-000 BAIRRO CIDADE - UF",
        _cpf(roteiro.cpf, "."),
    ]  # fmt: skip
    return "\n".join(linhas) + "\n"


# ----------------------------------------------------------------------------------- saídas
def textos(roteiro: Roteiro) -> list[tuple[str, str]]:
    """(nome do arquivo sem extensão, texto) de cada nota, na ordem do roteiro."""
    nomes = _nomes_de_pregao()
    saida: list[tuple[str, str]] = []
    for nota in roteiro.notas:
        texto = (texto_bmf(nota, roteiro) if nota.segmento == "bmf"
                 else texto_bovespa(nota, roteiro, nomes))  # fmt: skip
        nome = f"{roteiro.corretora.lower()}_{nota.data:%Y_%m_%d}_{nota.segmento}_{nota.numero}"
        saida.append((nome, texto))
    return saida


def pdf(texto: str, destino: Path) -> Path:
    """O texto num PDF de páginas A4 deitadas (folhas separadas por form feed), em fonte de
    largura fixa."""
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen.canvas import Canvas

    largura, altura = landscape(A4)
    tela = Canvas(str(destino), pagesize=(largura, altura))
    for pagina in texto.split("\f"):
        linhas = pagina.splitlines()
        tamanho = min(7.0, (altura - 48) / max(len(linhas), 1) / 1.15)
        tela.setFont("Courier", tamanho)
        y = altura - 28
        for linha in linhas:
            tela.drawString(28, y, linha)
            y -= tamanho * 1.15
        tela.showPage()
    tela.save()
    return destino


def movimentacao(roteiro: Roteiro, destino: Path) -> Path:
    """Os proventos do roteiro numa planilha de movimentação no formato da Área do Investidor."""
    from openpyxl import Workbook

    nomes = _nomes_de_pregao()
    livro = Workbook()
    folha = livro.active
    assert folha is not None
    folha.title = "Movimentação"
    folha.append(["Entrada/Saída", "Data", "Movimentação", "Produto", "Instituição",
                  "Quantidade", "Preço unitário", "Valor da Operação"])  # fmt: skip
    for dia, tipo, ticker, quantidade, unitario, valor in roteiro.proventos:
        produto = f"{ticker} - {nomes.get(ticker, ticker)}"
        folha.append(["Credito", f"{dia:%d/%m/%Y}", tipo, produto, INSTITUICAO_B3, quantidade,
                      float(unitario), float(valor)])  # fmt: skip
    livro.save(destino)
    return destino


def gerar(roteiro: Roteiro, pasta: Path) -> tuple[list[Path], Path | None]:
    """Grava os PDFs das notas e a planilha de proventos (se houver) na pasta."""
    pasta.mkdir(parents=True, exist_ok=True)
    pdfs = [pdf(texto, pasta / f"{nome}.pdf") for nome, texto in textos(roteiro)]
    planilha = None
    if roteiro.proventos:
        ano = roteiro.notas[-1].data.year
        planilha = movimentacao(roteiro, pasta / f"movimentacao-{ano}-demonstracao.xlsx")
    return pdfs, planilha


def main(argv: Sequence[str]) -> None:
    pdfs, planilha = gerar(carregar(Path(argv[0])), Path(argv[1]))
    print(f"{len(pdfs)} notas" + (f" e {planilha.name}" if planilha else ""))


if __name__ == "__main__":
    main(sys.argv[1:])
