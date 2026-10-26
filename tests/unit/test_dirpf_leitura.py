"""F8: leitura da declaração entregue (PDF "Imprimir > Declaração" do programa IRPF).

O texto abaixo é fictício, no layout do programa do exercício 2026 (cabeçalho por página, valores
em formato brasileiro, nomes que quebram a linha, valores fora do lugar). Os CNPJ são de companhias
abertas (dado público); o CPF é o 000.000.000-00 de exemplo.
"""

from decimal import Decimal

import pytest

from coir.erros import GabaritoError
from coir.gabarito.dirpf import FII, MESES, ler_texto

D = Decimal
CABECALHO = """NOME:
CPF: IMPOSTO SOBRE A RENDA - PESSOA FÍSICA
EXERCÍCIO 2026 ANO-CALENDÁRIO 2025
CONTRIBUINTE FICTICIO
DECLARAÇÃO DE AJUSTE ANUAL
000.000.000-00
"""
LINHAS_MERCADO = (
    "Mercado à Vista - Ações", "Mercado à Vista - Ouro",
    "Mercado à Vista - Ouro ativo financeiro fora de bolsa", "Mercado de Opções - Ações",
    "Mercado de Opções - Ouro", "Mercado de Opções - Fora de bolsa", "Mercado de Opções - Outros",
    "Mercado Futuro - Dólar dos EUA", "Mercado Futuro - Índices", "Mercado Futuro - Juros",
    "Mercado Futuro - Outros", "Mercado a Termo - Ações/Ouro", "Mercado a Termo - Outros",
)  # fmt: skip

RENDIMENTOS = """RENDIMENTOS ISENTOS E NÃO TRIBUTÁVEIS (Valores em Reais)
09. Lucros e dividendos recebidos 1.300,50
Beneficiário CPF CNPJ da Fonte Pagadora Nome da Fonte Pagadora Valor
Titular 000.000.000-00 33.592.510/0001-54 VALE S.A. 1.000,00
61.532.644/0001-15 ITAUSA
INVESTIMENTOS S.A.
Titular 000.000.000-00 300,50
12. Rendimentos de cadernetas de poupança, letras hipotecárias, letras de crédito do agronegócio e
certificados de recebíveis do agronegócio e imobiliários (CRA e CRI)
0,00
99. Outros
CPF/CNPJ da Fonte
Pagadora
Nome da Fonte
Pagadora
Beneficiário CPF Descrição Valor
Titular 000.000.000-00 33.592.510/0001-54 VALE S.A. REEMBOLSO DE
PROVENTOS BTC
12,34
TOTAL 1.312,84
RENDIMENTOS SUJEITOS À TRIBUTAÇÃO EXCLUSIVA / DEFINITIVA (Valores em Reais)
10. Juros sobre capital próprio 85,00
Beneficiário CPF CNPJ da Fonte Pagadora Nome da Fonte Pagadora Valor
Titular 000.000.000-00 61.532.644/0001-15 ITAUSA S.A. 85,00
TOTAL 85,00
RENDIMENTOS TRIBUTÁVEIS RECEBIDOS DE PESSOA JURÍDICA PELO TITULAR (IMPOSTO SUSPENSO)
Sem Informações
IMPOSTO PAGO / RETIDO (Valores em Reais)
01. Imposto complementar 0,00
02. Imposto pago no exterior pelo titular e pelos dependentes
03. Imposto sobre a renda na fonte (Lei 11.033/2004)
0,00
4,56
0,00
Imposto devido com os rendimentos no exterior
PAGAMENTOS EFETUADOS
Sem Informações
Página 1 de 3
"""

BENS_1 = """DECLARAÇÃO DE BENS E DIREITOS (Valores em Reais)
SITUAÇÃO EM
31/12/2024 31/12/2025
GRUPO CÓDIGO DISCRIMINAÇÃO
Possui perdas a compensar de acordo com a Lei nº 14.754, de 2023 (art. 9º)? Não
6.000,00 0,00
CNPJ: 33.592.510/0001-54
105 - BRASIL
Bem ou direito pertencente ao: Titular CPF: 000.000.000-00
Negociados em Bolsa: Sim Código de Negociação: VALE3
03 01 100 ACOES VALE3; CNPJ: 33.592.510/0001-54. PRECO MEDIO DE 60,00
EM 31/12/2024. Bem com usufruto: Não
Página 2 de 3
"""
BENS_2 = """DECLARAÇÃO DE BENS E DIREITOS (Valores em Reais)
SITUAÇÃO EM
31/12/2024 31/12/2025
GRUPO CÓDIGO DISCRIMINAÇÃO
0,00 24,00
105 - BRASIL
Bem ou direito pertencente ao: Titular CPF: 000.000.000-00
Negociados em Bolsa: Sim Código de Negociação: ITSAA100
04 04 OPÇÃO DE COMPRA - ITSAA100
Bem com usufruto: Não
TOTAL 6.000,00 24,00
DÍVIDAS E ÔNUS REAIS (Valores em Reais)
Sem Informações
"""


def _mes(mes: int) -> str:
    linhas = [f"GANHOS LÍQUIDOS OU PERDAS - {MESES[mes - 1]}"]
    for i, rotulo in enumerate(LINHAS_MERCADO):
        comum = "-1.234,56" if i == 0 and mes == 1 else ("100,00" if i == 0 else "0,00")
        linhas.append(f"{rotulo} {comum} {'-17,20' if i == 8 else '0,00'}")
    linhas += [
        "RESULTADO LÍQUIDO DO MÊS 0,00 0,00",
        "Resultado negativo até o mês anterior 1.000,00 0,00",
        "BASE DE CÁLCULO DO IMPOSTO 0,00 0,00",
        "Prejuízo a compensar 1.000,00 0,00",
        "Alíquota do imposto 15% 20%",
        "IMPOSTO DEVIDO 0,00 0,00",
        "Total do imposto devido 0,00",
        "IR fonte de day-trade do mês 0,00",
        "Imposto a pagar 0,00",
        "Imposto pago 0,00",
        "CONSOLIDAÇÃO DO MÊS",
        f"IR fonte (Lei nº 11.033/2004) a compensar {mes},00",
    ]
    return "\n".join(linhas)


def _renda_variavel(meses: range = range(1, 13)) -> str:
    return "\n".join([
        "RENDA VARIÁVEL - OPERAÇÕES COMUNS/DAYTRADE - TITULAR",
        "(Valores em Reais)",
        "TIPO DE MERCADO/ATIVO OPERAÇÕES COMUNS OPERAÇÕES DAY-TRADE",
        *(_mes(m) for m in meses),
        "RENDA VARIÁVEL - OPERAÇÕES COMUNS/DAYTRADE - DEPENDENTES",
        "Sem Informações",
        f"{FII} - TITULAR",
        "Sem Informações",
        f"{FII} - DEPENDENTES",
        "Sem Informações",
    ])  # fmt: skip


def _paginas(renda_variavel: str | None = None) -> list[str]:
    rv = _renda_variavel() if renda_variavel is None else renda_variavel
    return [CABECALHO + RENDIMENTOS, CABECALHO + BENS_1, CABECALHO + BENS_2 + rv]


def test_ficha_de_renda_variavel_mes_a_mes() -> None:
    d = ler_texto(_paginas())
    assert (d.exercicio, d.ano) == (2026, 2025)
    assert [m.mes for m in d.renda_variavel] == list(range(1, 13)) and d.fii == ()
    jan, dez = d.renda_variavel[0], d.renda_variavel[11]
    assert jan.valores["Mercado à Vista - Ações"] == (D("-1234.56"), D("0"))
    assert jan.valor("Mercado Futuro - Índices", 1) == D("-17.20")
    assert jan.valores["Alíquota do imposto"] == (D("0.15"), D("0.20"))
    assert jan.valores["Total do imposto devido"] == (D("0"), None)
    assert dez.valor("IR fonte (Lei nº 11.033/2004) a compensar") == D("12")
    with pytest.raises(GabaritoError, match="sem a coluna 1"):
        jan.valor("Imposto pago", 1)


def test_bens_com_quebra_de_pagina_e_usufruto_colado() -> None:
    bens = ler_texto(_paginas()).bens
    assert [(b.grupo, b.codigo, b.ticker, b.cnpj, b.anterior, b.atual) for b in bens] == [
        ("03", "01", "VALE3", "33592510000154", D("6000.00"), D("0")),
        ("04", "04", "ITSAA100", None, D("0"), D("24.00")),
    ]
    assert bens[0].discriminacao.endswith("EM 31/12/2024.")  # sem o "Bem com usufruto"


BENS_2024_1 = """DECLARAÇÃO DE BENS E DIREITOS (Valores em Reais)
GRUPO DISCRIMINAÇÃO SITUAÇÃO EM
31/12/2022 31/12/2023
CÓDIGO
03 ACOES DA EMPRESA VALE SA; CNPJ: 33.592.510/0001-54.
TOTALIZANDO 100 PAPEIS AO PRECO MEDIO DE 60,00.
0,00 6.000,00
CNPJ: 33.592.510/0001-54
01
105 - BRASIL
Bem ou direito pertencente ao: Titular CPF: 000.000.000-00
Negociados em Bolsa: Sim Código de Negociação: VALE3
99 JCP A RECEBER DA ITAUSA, CREDITADO EM 2023
Página 2 de 3
"""
BENS_2024_2 = """DECLARAÇÃO DE BENS E DIREITOS (Valores em Reais)
GRUPO DISCRIMINAÇÃO SITUAÇÃO EM
31/12/2022 31/12/2023
CÓDIGO
E PAGO EM 2024.
0,00 85,00
CNPJ: 61.532.644/0001-15
07
105 - BRASIL
Bem ou direito pertencente ao: Titular CPF: 000.000.000-00
TOTAL 0,00 6.085,00
DÍVIDAS E ÔNUS REAIS (Valores em Reais)
Sem Informações
"""


def test_bens_no_layout_ate_2024_com_quebra_de_pagina() -> None:
    """Exercício 2024: a discriminação vem antes dos valores e o código numa linha só."""
    paginas = [CABECALHO + RENDIMENTOS, CABECALHO + BENS_2024_1,
               CABECALHO + BENS_2024_2 + _renda_variavel()]  # fmt: skip
    bens = ler_texto(paginas).bens
    assert [(b.grupo, b.codigo, b.ticker, b.cnpj, b.anterior, b.atual) for b in bens] == [
        ("03", "01", "VALE3", "33592510000154", D("0"), D("6000.00")),
        ("99", "07", None, "61532644000115", D("0"), D("85.00")),
    ]
    assert bens[1].discriminacao == "JCP A RECEBER DA ITAUSA, CREDITADO EM 2023 E PAGO EM 2024."


def test_rendimentos_por_fonte_com_nome_quebrado() -> None:
    d = ler_texto(_paginas())
    assert d.totais[("isentos", "09")] == D("1300.50")
    assert [(r.codigo, r.cnpj, r.valor) for r in d.rendimentos] == [
        ("09", "33592510000154", D("1000.00")),
        ("09", "61532644000115", D("300.50")),
        ("99", "33592510000154", D("12.34")),
        ("10", "61532644000115", D("85.00")),
    ]
    assert d.totais[("isentos", "99")] == D("12.34")  # a linha 99 não imprime total
    assert d.totais[("isentos", "12")] == 0 and d.sem_fonte == ()  # total na linha de baixo
    assert d.ir_fonte_lei_11033 == D("4.56")


def test_fonte_ambigua_nao_e_chutada() -> None:
    paginas = _paginas()
    paginas[0] = paginas[0].replace("12,34\n", "12,34\n50,00\n")  # dois valores para uma fonte
    d = ler_texto(paginas)
    assert ("isentos", "99") in d.sem_fonte
    assert not any(r.codigo == "99" for r in d.rendimentos)
    assert d.totais[("isentos", "99")] == D("62.34")


def test_layout_inesperado_falha_alto() -> None:
    with pytest.raises(GabaritoError, match="sem os 12 meses"):
        ler_texto(_paginas(_renda_variavel(range(1, 12))))
    with pytest.raises(GabaritoError, match="cabeçalho"):
        ler_texto(["texto qualquer\n" * 8])
    sem_bens = _paginas()
    sem_bens[1] = CABECALHO + "Página 2 de 3\n"
    sem_bens[2] = CABECALHO + _renda_variavel()
    with pytest.raises(GabaritoError, match="BENS E DIREITOS"):
        ler_texto(sem_bens)


def test_layout_ate_2025_sem_bem_com_usufruto() -> None:
    """Até o exercício 2025 o programa não imprime "Bem com usufruto": o bem acaba quando começa
    o seguinte (linha de valores) ou no TOTAL; a pergunta do topo muda e se repete por página."""
    antiga = "Possui perdas a compensar de acordo com a Lei nº 14.754, de 2023 (art. 9º)? Não"
    nova = (
        "Atualizou o valor de algum bem imóvel e pagou o ganho de capital até 16/12/2024 de "
        "acordo com a Lei nº 14.973/2024? Não"
    )
    paginas = [p.replace(" Bem com usufruto: Não", "").replace("Bem com usufruto: Não\n", "")
               .replace(antiga, nova) for p in _paginas()]  # fmt: skip
    bens = ler_texto(paginas).bens
    assert [(b.ticker, b.anterior, b.atual) for b in bens] == [
        ("VALE3", D("6000.00"), D("0")), ("ITSAA100", D("0"), D("24.00"))]  # fmt: skip
    assert bens[0].discriminacao.endswith("EM 31/12/2024.")
