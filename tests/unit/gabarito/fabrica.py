"""Planilha-gabarito sintética em memória (mesma estrutura das abas reais, dados fictícios).

CPFs são montados em tempo de execução a partir de bases fictícias (o arquivo não contém nenhum
número que o verificador de dados pessoais bloquearia).
"""

from coir.gabarito.extrair import MESES

Linhas = list[list[object]]


def com_dv(base9: str) -> str:
    nums = [int(c) for c in base9]
    for pos in (9, 10):
        soma = sum(n * p for n, p in zip(nums, range(pos + 1, 1, -1), strict=False))
        nums.append((soma * 10) % 11 % 10)
    return "".join(map(str, nums))


def mascara(cpf: str) -> str:
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"


CPF_A = com_dv("123456789")
CPF_A_ERRADO = (
    CPF_A[:4] + str((int(CPF_A[4]) + 1) % 10) + CPF_A[5:]
)  # 1 dígito trocado, DV inválido
CONTA_A = 111222.0
CONTA_B = 333444.0
SERIAL_2008_04_18 = 39556.0
SERIAL_2008_05_20 = 39588.0

CAB_DADOS: list[object] = [
    "Corretora", "CPF", "Conta", "Nota", "Data", "C/V", "Papel", "Operacao", "Preço",
    "Quantidade", "Total", "Custos", "PM", "IRRF", "Prazo", "Exercício",
]  # fmt: skip
CAB_OP: list[object] = [
    None, "Nº", "Tipo", "Mercado", "Papel", "C/V", "Data", "Qtde", "PM", "PMA", "QtdeA",
    "Nota", "Custos", "IRRF", "Total", "Corretora", "Conta", "Mês", "Ano", "Isento",
]  # fmt: skip


CAB_CUSTODIA: list[object] = [
    "Corretora", "CPF", "Data", "C/V", "Código_Papel", "Operacao", "Preço_Médio", "Quantidade",
    "Total", "Prejuízo Acumulado",
]  # fmt: skip


def normais_dados() -> Linhas:
    return [
        CAB_DADOS,
        ["Agora", mascara(CPF_A), CONTA_A, 2588033.0, SERIAL_2008_04_18, "C", "GGBR4", "VISTA",
         62.2, 300.0, 18660.0, 26.919999999999998, 62.28973333333333, 0.0, None, None],
        ["XP ", mascara(CPF_A), CONTA_B, 11490605.0, SERIAL_2008_05_20, "V", "GGBR4", "VISTA",
         82.68, 100.0, 8268.0, 23.29, 82.4471, 0.41, None, None],
        ["XP", mascara(CPF_A_ERRADO), CONTA_B, None, SERIAL_2008_05_20, "C", "GGBR4",
         "Desdobramento", 0.0, 300.0, None, None, None, None, None, None],
        [None] * 16,
    ]  # fmt: skip


def daytrade_dados() -> Linhas:
    return [
        CAB_DADOS,
        ["XP", mascara(CPF_A), CONTA_B, 400729.0, SERIAL_2008_05_20, "C", "WING19", "INDICE",
         97310.0, 1.0, 19462.0, 0.34, 97310.0, 0.0, None, None],
    ]  # fmt: skip


def op(tipo: str, papel: str, cv: str, data: float, qtde: float, pma: float, qtde_a: float,
       nota: float | None, total: float, corretora: str, conta: float,
       competencia: tuple[float, float], isento: str | None = None) -> list[object]:  # fmt: skip
    mes, ano = competencia
    return [None, 1.0, tipo, tipo, papel, cv, data, qtde, pma, pma, qtde_a, nota, 1.0, 0.0, total,
            corretora, conta, mes, ano, isento]  # fmt: skip


def normais_op() -> Linhas:
    compra = op("VISTA", "GGBR4", "C", SERIAL_2008_04_18, 300, 62.2897, 300, 2588033, 0, "AGORA",
                CONTA_A, (4.0, 2008.0))  # fmt: skip
    # venda de maio apurada em junho (competência diferente da data do pregão), marcada isenta
    venda = op("VISTA", "GGBR4", "V", SERIAL_2008_05_20, 100, 62.2897, 200, 11490605, 2015.74, "XP",
               CONTA_B, (6.0, 2008.0), "Sim")  # fmt: skip
    custodia_vazia: list[object] = [None, 1.0, *([None] * 15)]  # Tipo vazio: ignorada
    return [CAB_OP, custodia_vazia, compra, venda]


def daytrade_op() -> Linhas:
    cab = [c for c in CAB_OP if c != "Mercado"]
    linha = op("INDICE", "WING19", "C", SERIAL_2008_05_20, 1, 97310, 1, 400729, 0, "XP", CONTA_B,
               (5.0, 2008.0))  # fmt: skip
    del linha[3]
    return [cab, linha]


def bloco_anual(ano: float, rotulo: str | None, valores_por_mes: dict[int, list[float]]) -> Linhas:
    cabeca: list[object] = (
        [None, None, None, rotulo, ano] if rotulo else [None, None, None, ano, "Tipo"]
    )
    linhas: Linhas = [cabeca, [None, None, None, "Resultado"]]
    for i, mes in enumerate(MESES, start=1):
        linhas.append([None, None, mes, *valores_por_mes.get(i, [0.0] * 11)])
    return linhas


def cont(resultado_maio: float) -> Linhas:
    maio = [resultado_maio, resultado_maio, 0, 0, 0, 0, 0, 0.41, -0.41, 0, 8268.0]
    linhas: Linhas = [[None]]
    return [*linhas, *bloco_anual(2008.0, None, {5: maio}), *bloco_anual(2009.0, None, {})]


def darf() -> Linhas:
    junho = [None, 39660.0, 39660.0, 6015.0, 302.36, None, None, 302.36]
    return bloco_anual(2008.0, "DARF", {6: junho})


def isentos() -> Linhas:
    linhas: Linhas = [
        [None, None, "Rendimentos"],
        [None, None, 2008.0, 2009.0, 2010.0, 2011.0, 2012.0],
    ]
    for mes in MESES:
        linhas.append([None, mes, 100.0 if mes == "Maio" else 0.0, 0.0, 0.0, 0.0, 0.0])
    return linhas


def eventos_b3() -> Linhas:
    return [
        ["HISTÓRICO", None],
        CAB_DADOS,
        ["XP", mascara(CPF_A), CONTA_B, None, SERIAL_2008_05_20, "C", "GGBR4", "Desdobramento",
         0.0, 300.0, None, None, None, None, None, None],
    ]  # fmt: skip


def planilha() -> dict[str, Linhas]:
    return {
        "Normais_Dados": normais_dados(),
        "DayTrade_Dados": daytrade_dados(),
        "Normais_Op": normais_op(),
        "DayTrade_Op": daytrade_op(),
        "Normais_Cont": cont(2015.74),
        "DayTrade_Cont": cont(0.0),
        "DARF": darf(),
        "Isentos": isentos(),
        "Eventos_B3": eventos_b3(),
        "Custodia_Inicial": [CAB_CUSTODIA],
        "BTC_Outros": bloco_anual(2008.0, "Outros Custos", {}),
    }
