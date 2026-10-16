"""Casos fiscais fechados à mão (especialista-fiscal-b3 + usuário, 14-15/10/2026).

Cada teste cita a regra e a fonte. Valores em reais; imposto truncado em 2 casas.
"""

from datetime import date
from decimal import Decimal

import pytest

from coir.domain.fiscal import Categoria, Grupo, TipoAjuste, TipoAtivo, TipoEvento, Visao
from coir.erros import MotorError
from tests.fiscal_cases.apoio import D, Livro, ajuste, categoria, mes, parametros


# ------------------------------------------------------------------------------------- T1 (R5)
def _t1() -> Livro:
    livro = Livro()
    livro.nota(date(2024, 2, 1), livro.op("C", "AAAA3", 1000, "16.20"),
               livro.op("C", "BBBB3", 1000, "28"))  # fmt: skip
    livro.nota(date(2024, 3, 5), livro.op("V", "AAAA3", 1000, "15"))  # vendas 15.000: isento
    livro.nota(date(2024, 3, 6), livro.op("C", "AAAAC10", 100, "5", mercado="CALL",
                                          vencimento=date(2024, 12, 20)))  # fmt: skip
    livro.nota(date(2024, 3, 7), livro.op("V", "AAAAC10", 100, "2", mercado="CALL",
                                          vencimento=date(2024, 12, 20)))  # fmt: skip
    livro.nota(date(2024, 4, 10), livro.op("V", "BBBB3", 1000, "30"), irrf="1.50")
    return livro


def test_t1_prejuizo_em_mes_isento_compensa() -> None:
    """R5: prejuízo de ações em mês isento continua compensável (Lei 8.981/1995 art. 72 §8º;
    Manual MIR 2026; decisão do usuário). Mar/24: ações -1.200 (vendas 15.000), opções -300.
    Abr/24: ganho 2.000 -> base 500, IR 75,00 - IRRF 1,50 = DARF 73,50 em 31/05/2024."""
    ap = _t1().apurar(date(2024, 4, 30))
    marco = categoria(ap, 2024, 3, Categoria.COMUM)
    assert (marco.vendas_acoes, marco.resultado, marco.prejuizo_a_compensar) == (
        D(15000), D(-1500), D(-1500))  # fmt: skip
    abril = categoria(ap, 2024, 4, Categoria.COMUM)
    assert (abril.base, abril.imposto) == (D(500), D("75.00"))
    m = mes(ap, 2024, 4)
    assert (m.irrf_deduzido, m.darf, m.vencimento) == (D("1.50"), D("73.50"), date(2024, 5, 31))


def test_t1_com_o_parametro_desligado_o_prejuizo_isento_se_perde() -> None:
    """R5, leitura alternativa (gabarito v2): prejuízo de mês isento descartado. Abr/24: base
    1.700, IR 255,00, DARF 253,50."""
    ap = _t1().apurar(date(2024, 4, 30), parametros(prejuizo_mes_isento_compensavel="false"))
    assert categoria(ap, 2024, 3, Categoria.COMUM).prejuizo_a_compensar == D(-300)
    abril = categoria(ap, 2024, 4, Categoria.COMUM)
    assert (abril.base, abril.imposto, mes(ap, 2024, 4).darf) == (
        D(1700), D("255.00"), D("253.50"))  # fmt: skip


# ------------------------------------------------------------------------------------- T2 (R4)
@pytest.mark.parametrize(("preco", "isento"), [("20", True), ("20.00001", False)])
def test_t2_limite_pelo_valor_bruto_das_vendas(preco: str, isento: bool) -> None:
    """R4: o limite de R$ 20 mil é sobre o valor bruto das alienações (Lei 11.033/2004 art. 3º
    I); custos de venda não reduzem o valor para o limite. 20.000,00 com custos 10 é isento;
    20.000,01 é tributável."""
    livro = Livro()
    livro.nota(date(2024, 1, 2), livro.op("C", "AAAA3", 1000, "15"))
    livro.nota(date(2024, 2, 1), livro.op("V", "AAAA3", 1000, preco), custos="10")
    c = categoria(livro.apurar(date(2024, 2, 29)), 2024, 2, Categoria.COMUM)
    assert c.vendas_acoes == (D("20000.00") if isento else D("20000.01"))
    assert (c.isento > 0) is isento and (c.base > 0) is not isento


def test_t2_units_fora_do_limite_e_tributadas() -> None:
    """R4a: units não entram no limite nem na isenção (decisão do usuário; parâmetro
    units_na_isencao). Vendas de ações 16.000 (isentas) + units 5.000 (ganho tributado)."""
    livro = Livro()
    livro.nota(date(2024, 1, 2), livro.op("C", "AAAA3", 1000, "15"),
               livro.op("C", "TAEE11", 100, "40"))  # fmt: skip
    livro.nota(date(2024, 2, 1), livro.op("V", "AAAA3", 1000, "16"),
               livro.op("V", "TAEE11", 100, "50"))  # fmt: skip
    c = categoria(livro.apurar(date(2024, 2, 29)), 2024, 2, Categoria.COMUM)
    assert (c.vendas_acoes, c.isento, c.resultado, c.imposto) == (
        D(16000), D(1000), D(1000), D("150.00"))  # fmt: skip


# ------------------------------------------------------------------------------------- T3 (R6)
def test_t3_irrf_comum_e_day_trade_abatem_o_imposto_total() -> None:
    """R6: IRRF de 0,005% e de 1% abatem o imposto do mês (saldo único). DT ganho 1.000:
    IR 200 - 10 (1%) - 2 (0,005%) = 188."""
    livro = Livro()
    livro.nota(date(2024, 5, 2), livro.op("C", "AAAA3", 1000, "10"),
               livro.op("V", "AAAA3", 1000, "11"), irrf="2", irrf_dt="10")  # fmt: skip
    ap = livro.apurar(date(2024, 5, 31))
    assert categoria(ap, 2024, 5, Categoria.DAY_TRADE).imposto == D("200.00")
    assert mes(ap, 2024, 5).darf == D("188.00")


# ------------------------------------------------------------------------------------ T4 (R12)
def test_t4_darf_abaixo_de_10_acumula_e_vira_o_ano() -> None:
    """R12: DARF abaixo de R$ 10 acumula, inclusive para o ano seguinte (Lei 9.430/1996 art. 68
    §1º); vence no último dia útil do mês seguinte (RIR/2018 art. 915). Dez/23 7,00 + Jan/24
    4,00 = 11,00, vencimento 29/02/2024."""
    livro = Livro()
    livro.nota(date(2023, 11, 1), livro.op("C", "TAEE11", 200, "10"))
    livro.nota(date(2023, 12, 1), livro.op("V", "TAEE11", 100, "10.4667"))  # ganho 46,67
    livro.nota(date(2024, 1, 2), livro.op("V", "TAEE11", 100, "10.2667"))  # ganho 26,67
    ap = livro.apurar(date(2024, 1, 31))
    dez, jan = mes(ap, 2023, 12), mes(ap, 2024, 1)
    assert (dez.imposto_devido, dez.darf, dez.darf_acumulado) == (D("7.00"), 0, D("7.00"))
    assert (jan.imposto_devido, jan.darf, jan.vencimento) == (
        D("4.00"), D("11.00"), date(2024, 2, 29))  # fmt: skip


# ------------------------------------------------------------------------------------- T5 (R9)
def _exercicio(serie_cv: str, premio: str, exerc: str, cv: str, strike: str) -> Livro:
    livro = Livro()
    mercado = "CALL" if exerc == "EXERC CALL" else "PUT"
    livro.nota(date(2024, 1, 10), livro.op(serie_cv, "AAAAX30", 1000, premio, mercado=mercado,
                                           vencimento=date(2024, 2, 16)), custos="5")  # fmt: skip
    return livro


def test_t5_titular_de_call_soma_o_premio_ao_custo() -> None:
    """R9 (IN RFB 1.585/2015 art. 58): prêmio 1.000 + custo 5; exercício 1.000 a 30 com custos
    20 -> custo 31.025, PM 31,025."""
    livro = _exercicio("C", "1", "EXERC CALL", "C", "30")
    livro.nota(date(2024, 2, 16), livro.op("C", "AAAA3", 1000, "30", mercado="EXERC CALL",
                                           opcao="AAAAX30"), custos="20")  # fmt: skip
    pos = [m for m in livro.movimentos(date(2024, 2, 29)) if m.ticker == "AAAA3"][-1]
    assert (pos.quantidade_apos, pos.custo_apos, pos.pm_apos) == (1000, D(31025), D("31.025"))


def test_t5_lancador_de_call_soma_o_premio_a_alienacao() -> None:
    """R9: prêmio 1.000 - custo 5 = 995; exercício vende 1.000 a 30 com custos 20 -> alienação
    30.975; com PM 25, ganho 5.975 (comum; exercício não é day trade nem tem IRRF de DT)."""
    livro = _exercicio("V", "1", "EXERC CALL", "V", "30")
    livro.nota(date(2024, 1, 2), livro.op("C", "AAAA3", 1000, "25"))
    livro.nota(date(2024, 2, 16), livro.op("V", "AAAA3", 1000, "30", mercado="EXERC CALL",
                                           opcao="AAAAX30"), custos="20")  # fmt: skip
    venda = [m for m in livro.movimentos(date(2024, 2, 29)) if m.ticker == "AAAA3"][-1]
    assert (venda.valor, venda.resultado, venda.categoria) == (
        D(30975), D(5975), Categoria.COMUM)  # fmt: skip


def test_t5_titular_de_put_subtrai_o_premio_da_alienacao() -> None:
    """R9: prêmio 0,80 x 1.000 + custo 5 = 805; exercício vende 1.000 a 28 com custos 20 ->
    alienação 27.175."""
    livro = _exercicio("C", "0.80", "EXERC PUT", "V", "28")
    livro.nota(date(2024, 1, 2), livro.op("C", "AAAA3", 1000, "25"))
    livro.nota(date(2024, 2, 16), livro.op("V", "AAAA3", 1000, "28", mercado="EXERC PUT",
                                           opcao="AAAAX30"), custos="20")  # fmt: skip
    venda = [m for m in livro.movimentos(date(2024, 2, 29)) if m.ticker == "AAAA3"][-1]
    assert venda.valor == D(27175)


def test_t5_lancador_de_put_subtrai_o_premio_do_custo() -> None:
    """R9: prêmio 0,80 x 1.000 - custo 5 = 795; exercício compra 1.000 a 28 com custos 20 ->
    custo 27.225, PM 27,225."""
    livro = _exercicio("V", "0.80", "EXERC PUT", "C", "28")
    livro.nota(date(2024, 2, 16), livro.op("C", "AAAA3", 1000, "28", mercado="EXERC PUT",
                                           opcao="AAAAX30"), custos="20")  # fmt: skip
    pos = [m for m in livro.movimentos(date(2024, 2, 29)) if m.ticker == "AAAA3"][-1]
    assert (pos.custo_apos, pos.pm_apos) == (D(27225), D("27.225"))


def test_t5_ganho_da_venda_por_exercicio_nao_e_isento() -> None:
    """IN RFB 1.585/2015 art. 59 §2º IV: a isenção não se aplica à alienação de ações em
    exercício de opções. Lançador de 600 calls (prêmio 600) exercido: vende 600 a 30 (18.000,
    abaixo do limite) com PM 25 -> ganho 3.600 tributável, IR 540,00, DARF em 28/03/2024."""
    livro = Livro()
    livro.nota(date(2024, 1, 2), livro.op("C", "AAAA3", 600, "25"))
    livro.nota(date(2024, 1, 10), livro.op("V", "AAAAB30", 600, "1", mercado="CALL",
                                           vencimento=date(2024, 2, 16)))  # fmt: skip
    livro.nota(date(2024, 2, 16), livro.op("V", "AAAA3", 600, "30", mercado="EXERC CALL",
                                           opcao="AAAAB30"))  # fmt: skip
    ap = livro.apurar(date(2024, 2, 29))
    c = categoria(ap, 2024, 2, Categoria.COMUM)
    assert (c.isento, c.base, c.imposto) == (0, D(3600), D("540.00"))
    assert mes(ap, 2024, 2).vencimento == date(2024, 3, 28)


def test_t5_exercicio_sem_a_opcao_em_carteira_falha_alto() -> None:
    livro = Livro()
    livro.nota(date(2024, 2, 16), livro.op("C", "AAAA3", 1000, "30", mercado="EXERC CALL",
                                           opcao="AAAAX30"))  # fmt: skip
    with pytest.raises(MotorError, match="posição comprada"):
        livro.movimentos(date(2024, 2, 29))


# ------------------------------------------------------------------------------------ T6 (R10)
@pytest.mark.parametrize(("cv", "resultado"), [("C", "-505"), ("V", "495")])
def test_t6_opcao_que_vira_po_no_vencimento(cv: str, resultado: str) -> None:
    """R10: 1.000 opções a 0,50 com custo 5, não exercidas: titular perde 505; lançador ganha
    495, no mês do vencimento (IN RFB 1.585/2015)."""
    livro = Livro()
    livro.nota(date(2024, 1, 10), livro.op(cv, "AAAAB50", 1000, "0.50", mercado="CALL",
                                           vencimento=date(2024, 2, 16)), custos="5")  # fmt: skip
    ap = livro.apurar(date(2024, 2, 29))
    assert categoria(ap, 2024, 1, Categoria.COMUM).resultado == 0
    fev = categoria(ap, 2024, 2, Categoria.COMUM)
    assert fev.por_grupo[Grupo.OPCOES] == D(resultado)


# ------------------------------------------------------------------------------------- T7 (R2)
def test_t7_day_trade_pela_ordem_dos_negocios_e_excedente_comum() -> None:
    """R2: IN RFB 1.585/2015 art. 65 §3º ("pela ordem, o 1º negócio de compra com o 1º de
    venda") e §2º (o estoque anterior não entra). Posição 100 a PM 10; no dia, na ordem da nota:
    compra 100 a 11, compra 200 a 13, venda 200 a 14. Day trade: 100 a 11 + 100 a 13 = 2.400
    contra 2.800 -> ganho 400; excedente comum: 100 a 13 -> PM (1.000 + 1.300) / 200 = 11,50."""
    livro = Livro()
    livro.nota(date(2024, 3, 1), livro.op("C", "AAAA3", 100, "10"))
    livro.nota(date(2024, 3, 4), livro.op("C", "AAAA3", 100, "11"),
               livro.op("C", "AAAA3", 200, "13"), livro.op("V", "AAAA3", 200, "14"))  # fmt: skip
    ap = livro.apurar(date(2024, 3, 31))
    assert categoria(ap, 2024, 3, Categoria.DAY_TRADE).resultado == D("400.00")
    ultimo = [m for m in ap.movimentos if m.ticker == "AAAA3"][-1]
    assert (ultimo.quantidade_apos, ultimo.pm_apos) == (200, D("11.5"))


def test_t7_ordem_entre_notas_do_mesmo_dia_pelo_numero_da_nota() -> None:
    """Duas notas no mesmo dia e corretora: a de número menor vem primeiro (art. 65 §3º)."""
    livro = Livro()
    n2 = livro.nota(date(2024, 3, 4), livro.op("C", "AAAA3", 100, "13"))
    n1 = livro.nota(date(2024, 3, 4), livro.op("C", "AAAA3", 100, "11"),
                    livro.op("V", "AAAA3", 100, "14"))  # fmt: skip
    livro.notas = [n2.model_copy(update={"numero": "20"}), n1.model_copy(update={"numero": "10"})]
    ap = livro.apurar(date(2024, 3, 31))
    assert categoria(ap, 2024, 3, Categoria.DAY_TRADE).resultado == D(300)  # 1.400 - 1.100


def test_t7_day_trade_em_corretoras_diferentes_so_conta_antes_de_2010() -> None:
    """MP 497/2010: desde 28/07/2010 o day trade é na mesma corretora; antes, em qualquer."""
    for dia, dt in ((date(2010, 7, 27), True), (date(2010, 7, 28), False)):
        livro = Livro()
        livro.nota(dia, livro.op("C", "AAAA3", 100, "10"), corretora="XP")
        livro.nota(dia, livro.op("V", "AAAA3", 100, "11"), corretora="AGORA", conta=2)
        ap = livro.apurar(dia)
        assert (categoria(ap, dia.year, dia.month, Categoria.DAY_TRADE).resultado == 100) is dt


def test_t7_exercicio_e_venda_no_mesmo_dia_nao_e_day_trade() -> None:
    """IN RFB 1.585/2015 art. 65 §13: exercício de opção não caracteriza day trade."""
    livro = Livro()
    livro.nota(date(2024, 1, 10), livro.op("C", "AAAAX30", 1000, "1", mercado="CALL",
                                           vencimento=date(2024, 2, 16)))  # fmt: skip
    livro.nota(date(2024, 2, 16), livro.op("C", "AAAA3", 1000, "30", mercado="EXERC CALL",
                                           opcao="AAAAX30"),
               livro.op("V", "AAAA3", 1000, "33"))  # fmt: skip
    ap = livro.apurar(date(2024, 2, 29))
    assert categoria(ap, 2024, 2, Categoria.DAY_TRADE).resultado == 0
    assert categoria(ap, 2024, 2, Categoria.COMUM).resultado == D(2000)  # 33.000 - 31.000


# ------------------------------------------------------------------------------------- T8 (R3)
def test_t8_day_trade_de_fii_fica_na_categoria_fii() -> None:
    """R3: FII (inclusive day trade) a 20%, compensação só entre FII (IN RFB 1.585/2015).
    Ganho 100: IR 20,00 - IRRF 1,00 = 19,00."""
    livro = Livro()
    livro.nota(date(2024, 6, 3), livro.op("C", "XPML11", 100, "100", tipo=TipoAtivo.FII),
               livro.op("V", "XPML11", 100, "101", tipo=TipoAtivo.FII),
               irrf_dt="1")  # fmt: skip
    ap = livro.apurar(date(2024, 6, 30))
    fii = categoria(ap, 2024, 6, Categoria.FII)
    assert (fii.resultado, fii.imposto) == (D(100), D("20.00"))
    assert categoria(ap, 2024, 6, Categoria.DAY_TRADE).resultado == 0
    assert mes(ap, 2024, 6).darf == D("19.00")


# ------------------------------------------------------------------------------------ T9 (R18)
def test_t9_termo_entra_na_carteira_na_liquidacao() -> None:
    """R18 (decisão do usuário, 14/10/2026): posição 1.000 a PM 25; compra a termo de 1.000 a
    31,00 com custos 40. Venda de 500 antes da liquidação usa PM 25; depois da liquidação,
    PM (12.500 + 31.040) / 1.500 = 29,02666667."""
    livro = Livro()
    livro.nota(date(2024, 1, 2), livro.op("C", "AAAA3", 1000, "25"))
    livro.nota(date(2024, 1, 3), livro.op("C", "AAAA3", 1000, "31", mercado="TERMO",
                                          liquidacao=date(2024, 3, 4)), custos="40")  # fmt: skip
    livro.nota(date(2024, 2, 1), livro.op("V", "AAAA3", 500, "30"))
    movs = [m for m in livro.movimentos(date(2024, 3, 31)) if m.ticker == "AAAA3"]
    venda = next(m for m in movs if m.cv == "V")
    assert venda.custo_baixa == D(12500)
    assert movs[-1].pm_apos.quantize(D("0.00000001")) == D("29.02666667")


def test_termo_liquidado_depois_da_venda_das_acoes_realiza_o_resultado() -> None:
    """R18 + venda a descoberto (leitura A1, decisão do usuário de 15/10/2026; IN RFB 1.585/2015
    art. 58: custo médio): sem posição, compra a termo de 1.000 a 31,00 (custos 40) em 03/01 e
    venda das 1.000 a 35,00 em 01/02, antes da liquidação (04/03). A liquidação cobre a posição
    vendida: resultado 35.000 - 31.040 = 3.960 em 03/2024, com o limite de isenção de 02/2024
    (vendas 35.000 > 20.000: tributável, 15% = 594,00)."""
    livro = Livro()
    livro.nota(date(2024, 1, 3), livro.op("C", "AAAA3", 1000, "31", mercado="TERMO",
                                          liquidacao=date(2024, 3, 4)), custos="40")  # fmt: skip
    livro.nota(date(2024, 2, 1), livro.op("V", "AAAA3", 1000, "35"))
    ap = livro.apurar(date(2024, 3, 31))
    termo = next(m for m in ap.movimentos if m.origem.startswith("termo:"))
    assert (termo.resultado, termo.quantidade_apos, termo.mes_limite) == (
        D(3960), D(0), date(2024, 2, 1))  # fmt: skip
    assert categoria(ap, 2024, 2, Categoria.COMUM).resultado == 0
    assert categoria(ap, 2024, 3, Categoria.COMUM).imposto == D("594.00")


def test_ganho_na_liquidacao_do_termo_nunca_e_isento() -> None:
    """IN RFB 1.585/2015 art. 59 §2º IV: a isenção de R$ 20 mil não se aplica à alienação no
    vencimento ou na liquidação antecipada do termo. Compra a termo de 500 a 31,00 (custos 20) em
    03/01/2024 e venda das 500 a 35,00 (17.500, abaixo do limite) em 01/02/2024, antes da
    liquidação (04/03): resultado 17.500 - 15.520 = 1.980 em 03/2024, tributável (15% = 297,00)."""
    livro = Livro()
    livro.nota(date(2024, 1, 3), livro.op("C", "AAAA3", 500, "31", mercado="TERMO",
                                          liquidacao=date(2024, 3, 4)), custos="20")  # fmt: skip
    livro.nota(date(2024, 2, 1), livro.op("V", "AAAA3", 500, "35"))
    ap = livro.apurar(date(2024, 3, 31))
    marco = categoria(ap, 2024, 3, Categoria.COMUM)
    assert (marco.resultado, marco.isento, marco.imposto) == (D(1980), D(0), D("297.00"))


# --------------------------------------------------------------------------------------- outros
def test_compensacao_separada_entre_comum_e_day_trade() -> None:
    """R2: prejuízo de day trade só compensa com ganho de day trade (IN RFB 1.585/2015)."""
    livro = Livro()
    livro.nota(date(2024, 1, 3), livro.op("C", "AAAA3", 1000, "10"),
               livro.op("V", "AAAA3", 1000, "9"))  # DT -1.000  # fmt: skip
    livro.nota(date(2024, 2, 1), livro.op("C", "TAEE11", 100, "10"))
    livro.nota(date(2024, 2, 5), livro.op("V", "TAEE11", 100, "20"))  # comum +1.000
    ap = livro.apurar(date(2024, 2, 29))
    assert categoria(ap, 2024, 2, Categoria.COMUM).imposto == D("150.00")
    assert categoria(ap, 2024, 2, Categoria.DAY_TRADE).prejuizo_a_compensar == D(-1000)


def test_saldo_de_irrf_passa_no_ano_e_dezembro_vai_para_a_declaracao() -> None:
    """R6: saldo de IRRF transporta aos meses seguintes do ano e não passa para janeiro. O de
    0,005% vai para a declaração (IN 1.585 art. 63 §8º III); o de day trade só restituição
    (art. 65 §§9º e 12 II; P&R 715)."""
    livro = Livro()
    livro.nota(date(2023, 11, 1), livro.op("C", "TAEE11", 100, "10"), irrf="3")
    livro.nota(date(2023, 12, 1), livro.op("C", "AAAA3", 10, "10"), irrf_dt="2")
    livro.nota(date(2024, 1, 2), livro.op("V", "TAEE11", 100, "11"))  # ganho 100: IR 15
    ap = livro.apurar(date(2024, 1, 31))
    assert mes(ap, 2023, 12).irrf_saldo == D(5)
    [anual] = ap.irrf_anual
    assert (anual.ano, anual.saldo, anual.comum, anual.day_trade) == (2023, D(5), D(3), D(2))
    assert (anual.saldo_comum, anual.saldo_day_trade) == (D(3), D(2))
    assert (mes(ap, 2024, 1).irrf_saldo_anterior, mes(ap, 2024, 1).darf) == (0, D("15.00"))


def test_futuros_resultado_no_encerramento_da_posicao() -> None:
    """R11: Lei 11.051/2004 art. 32; IN 1.585 art. 61; P&R 728 (resultado na liquidação ou
    encerramento). Compra de 1 WINM24 em 25/03/2024; ajustes de 25 a 28/03 somam +300 (100 no
    negócio + 200 de AJUPOS); de 01 a 05/04 somam -100 (AJUPOS -50 + ajuste do negócio de
    encerramento -50). Março: 0. Abril: +200, IR 30,00, DARF em 31/05/2024."""
    livro = Livro()
    livro.nota(date(2024, 3, 25), ajustes=[
        ajuste("WINM24", "100", cv="C", quantidade="1", volume="26000")])  # fmt: skip
    livro.nota(date(2024, 3, 28), ajustes=[ajuste("WINM24", "200", tipo=TipoAjuste.AJUPOS)])
    livro.nota(date(2024, 4, 1), ajustes=[ajuste("WINM24", "-50", tipo=TipoAjuste.AJUPOS)])
    livro.nota(date(2024, 4, 5), ajustes=[
        ajuste("WINM24", "-50", cv="V", quantidade="1", volume="26200")])  # fmt: skip
    ap = livro.apurar(date(2024, 4, 30))
    assert categoria(ap, 2024, 3, Categoria.COMUM).resultado == 0
    abril = categoria(ap, 2024, 4, Categoria.COMUM)
    assert (abril.por_grupo[Grupo.INDICE], abril.imposto) == (D(200), D("30.00"))
    assert mes(ap, 2024, 4).vencimento == date(2024, 5, 31)


def test_futuros_encerramento_parcial_pelo_preco_medio_e_vencimento() -> None:
    """Encerramento parcial: sai a parte proporcional dos ajustes acumulados (preço médio,
    decisão do usuário); no vencimento sai o resto. Custos rateados pelo volume; ajuste de day
    trade vai para DAY_TRADE (IN 1.585 art. 65 §§10 e 11)."""
    livro = Livro()
    venc = date(2024, 6, 12)
    livro.nota(date(2024, 5, 2), custos="4", ajustes=[
        ajuste("WINM24", "40", cv="C", quantidade="2", volume="100", vencimento=venc),
        ajuste("WINM24", "50", tipo=TipoAjuste.DAY_TRADE, volume="100", vencimento=venc),
    ])  # fmt: skip
    livro.nota(date(2024, 5, 3), ajustes=[
        ajuste("WINM24", "60", tipo=TipoAjuste.AJUPOS, vencimento=venc),
        ajuste("WINM24", "10", cv="V", quantidade="1", volume="0", vencimento=venc),
    ])  # fmt: skip
    livro.nota(date(2024, 6, 12), ajustes=[
        ajuste("WINM24", "-20", tipo=TipoAjuste.AJUPOS, vencimento=venc)])  # fmt: skip
    ap = livro.apurar(date(2024, 6, 30))
    maio = categoria(ap, 2024, 5, Categoria.COMUM)
    assert maio.por_grupo[Grupo.INDICE] == D(59)  # (40 - 2 + 60) / 2 + 10
    assert categoria(ap, 2024, 5, Categoria.DAY_TRADE).por_grupo[Grupo.INDICE] == D(48)
    assert categoria(ap, 2024, 6, Categoria.COMUM).por_grupo[Grupo.INDICE] == D(29)  # 49 - 20


def test_futuros_inversao_de_posicao_e_falhas_altas() -> None:
    """Lei 11.051/2004 art. 32: comprado em 1, a venda de 3 encerra 1 (resultado: ajuste
    acumulado 10 + 1/3 do ajuste do negócio 30 = 20) e abre vendido em 2, que sai no vencimento
    (20). Regra 5 (falhar alto): custos de nota BM&F sem nenhuma linha para ratear dão erro; posição
    sem vencimento conhecido que não encerra gera aviso."""
    livro = Livro()
    venc = date(2024, 6, 12)
    livro.nota(date(2024, 5, 2), ajustes=[
        ajuste("WINM24", "10", cv="C", quantidade="1", volume="100", vencimento=venc)])  # fmt: skip
    livro.nota(date(2024, 5, 3), ajustes=[
        ajuste("WINM24", "30", cv="V", quantidade="3", volume="300", vencimento=venc)])  # fmt: skip
    livro.nota(date(2024, 5, 6), ajustes=[ajuste("WDOM24", "5", cv="C", quantidade="1",
                                                 volume="50", grupo=Grupo.DOLAR)])  # fmt: skip
    ap = livro.apurar(date(2024, 6, 30))
    assert categoria(ap, 2024, 5, Categoria.COMUM).por_grupo[Grupo.INDICE] == D(20)
    assert categoria(ap, 2024, 6, Categoria.COMUM).por_grupo[Grupo.INDICE] == D(20)
    assert any("WDOM24" in a and "sem vencimento" in a for a in ap.avisos)
    livro.nota(date(2024, 6, 20), custos="3")
    with pytest.raises(MotorError, match="sem negócio nem ajuste"):
        livro.apurar(date(2024, 6, 30))


def test_eventos_ajustam_quantidade_e_custo_sem_resultado() -> None:
    """R8: bonificação soma quantidade com custo atribuído; desdobramento e grupamento só a
    quantidade; redução de capital reduz o custo (RIR/2018; Perguntas e Respostas IRPF)."""
    livro = Livro()
    livro.nota(date(2024, 1, 2), livro.op("C", "AAAA3", 100, "10"))
    livro.evento(date(2024, 2, 1), "AAAA3", TipoEvento.DESDOBRAMENTO, "C", "100")
    livro.evento(date(2024, 3, 1), "AAAA3", TipoEvento.BONIFICACAO, "C", "20", "100")
    livro.evento(date(2024, 4, 1), "AAAA3", TipoEvento.GRUPAMENTO, "V", "110")
    livro.evento(date(2024, 5, 1), "AAAA3", TipoEvento.REDUCAO_CAPITAL, "V", None, "300")
    movs = livro.movimentos(date(2024, 5, 31))
    ultimo = movs[-1]
    assert (ultimo.quantidade_apos, ultimo.custo_apos) == (110, D(800))
    assert all(m.resultado == 0 for m in movs)


def test_transferencia_de_custodia_so_muda_a_visao_por_conta() -> None:
    """Transferência entre contas do mesmo titular (ex.: de uma corretora para outra): leva a
    quantidade e o PM; sem resultado; o PM do titular não muda."""
    livro = Livro()
    livro.nota(date(2024, 1, 4), livro.op("C", "AAAA3", 1000, "8"), conta=1, corretora="AGORA")
    livro.evento(date(2024, 2, 17), "AAAA3", TipoEvento.TRANSFERENCIA, quantidade="700",
                 conta=1, destino=2, corretora="AGORA", corretora_destino="RICO")  # fmt: skip
    livro.nota(date(2024, 2, 17), livro.op("V", "AAAA3", 700, "9"), conta=2, corretora="RICO")
    por_conta = livro.movimentos(date(2024, 3, 31), Visao.CONTA)
    venda = next(m for m in por_conta if m.cv == "V" and m.origem.startswith("op:"))
    assert (venda.conta, venda.resultado, venda.quantidade_apos) == (2, D(700), 0)
    global_ = livro.movimentos(date(2024, 3, 31))
    assert [m.resultado for m in global_ if m.cv == "V"] == [D(700)]
    assert global_[-1].quantidade_apos == 300


# -------------------------------------------------------------- PM fiscal x visões operacionais
def _tres_contas() -> Livro:
    livro = Livro()
    livro.nota(date(2023, 1, 2), livro.op("C", "AAAA3", 100, "10"), conta=1, corretora="XP")
    livro.nota(date(2023, 1, 3), livro.op("C", "AAAA3", 100, "20"), conta=2, corretora="XP")
    livro.nota(date(2023, 1, 4), livro.op("C", "AAAA3", 200, "30"), conta=3, corretora="BTG")
    livro.nota(date(2023, 2, 1), livro.op("V", "AAAA3", 100, "25"), conta=1, corretora="XP")
    return livro


def test_pm_fiscal_unico_por_titular_e_a_conta_de_origem_no_movimento() -> None:
    """IN RFB 1.585/2015 art. 58: custo médio de aquisição do contribuinte, sem separar por
    corretora (decisões da F1 e do usuário, 15/10/2026). Compras: 100 a 10 (XP conta 1), 100 a 20
    (XP conta 2), 200 a 30 (BTG): PM (1.000 + 2.000 + 6.000) / 400 = 22,50. Venda de 100 a 25 na
    conta 1: ganho 2.500 - 2.250 = 250 (isento: vendas 2.500). O movimento guarda a conta e a
    corretora de origem; a posição que sobra é 300 a 22,50."""
    livro = _tres_contas()
    venda = next(m for m in livro.apurar(date(2023, 2, 28)).movimentos if m.cv == "V")
    assert (venda.resultado, venda.conta, venda.corretora) == (D(250), 1, "XP")
    assert categoria(livro.apurar(date(2023, 2, 28)), 2023, 2, Categoria.COMUM).isento == D(250)
    [pos] = livro.carteira(date(2023, 2, 28))
    assert (pos.corretora, pos.conta, pos.quantidade, pos.pm) == (None, None, 300, D("22.5"))


def test_carteira_por_corretora_e_por_conta_e_so_visao_operacional() -> None:
    """A mesma posição vista por corretora (XP: 200 a 15 -> vende 100 -> 100 a 15; BTG: 200 a 30)
    e por conta (conta 1 zerada; conta 2: 100 a 20; conta 3: 200 a 30). A quantidade fecha com a
    fiscal (300); o custo não, por desenho. A apuração usa só a visão por titular."""
    livro = _tres_contas()
    dia = date(2023, 2, 28)
    por_corretora = {
        p.corretora: (p.quantidade, p.pm) for p in livro.carteira(dia, Visao.CORRETORA)
    }
    assert por_corretora == {"BTG": (200, D(30)), "XP": (100, D(15))}
    por_conta = {(p.corretora, p.conta): (p.quantidade, p.pm)
                 for p in livro.carteira(dia, Visao.CONTA)}  # fmt: skip
    assert por_conta == {("BTG", 3): (200, D(30)), ("XP", 2): (100, D(20))}
    venda_xp = next(m for m in livro.movimentos(dia, Visao.CORRETORA) if m.cv == "V")
    assert venda_xp.resultado == D(1000)  # 2.500 - 100 x 15: não é o resultado fiscal


def test_transferencia_entre_contas_da_mesma_corretora_nao_muda_a_visao_por_corretora() -> None:
    """Transferência XP conta 1 -> XP conta 2: por corretora e por titular a posição é a mesma;
    por conta, ela passa para a conta 2 com o mesmo PM."""
    livro = Livro()
    livro.nota(date(2023, 1, 2), livro.op("C", "AAAA3", 100, "10"), conta=1, corretora="XP")
    livro.evento(date(2023, 1, 10), "AAAA3", TipoEvento.TRANSFERENCIA, conta=1, destino=2)
    dia = date(2023, 1, 31)
    assert [(p.corretora, p.conta, p.quantidade) for p in livro.carteira(dia, Visao.CORRETORA)] == [
        ("XP", None, 100)]  # fmt: skip
    assert [(p.conta, p.quantidade, p.pm) for p in livro.carteira(dia, Visao.CONTA)] == [
        (2, 100, D(10))]  # fmt: skip


def test_day_trade_entre_corretoras_antes_de_2010_fica_sem_conta_de_origem() -> None:
    """MP 497/2010: antes de 28/07/2010 o day trade junta corretoras; o movimento fica sem conta
    e sem corretora (a origem lista as duas operações). O pó de uma série comprada numa conta só
    fica nessa conta."""
    livro = Livro()
    livro.nota(date(2009, 3, 2), livro.op("C", "AAAA3", 100, "10"), conta=1, corretora="XP")
    livro.nota(date(2009, 3, 2), livro.op("V", "AAAA3", 100, "11"), conta=3, corretora="BTG")
    livro.nota(
        date(2009, 3, 2),
        livro.op("C", "AAAAC10", 100, "1", mercado="CALL", vencimento=date(2009, 3, 16)),
        conta=3,
        corretora="BTG",
    )
    movs = livro.movimentos(date(2009, 3, 31))
    dt = next(m for m in movs if m.categoria is Categoria.DAY_TRADE)
    assert (dt.resultado, dt.conta, dt.corretora) == (D(100), None, None)
    po = next(m for m in movs if m.origem == "vencimento")
    assert (po.conta, po.corretora, po.resultado) == (3, "BTG", D(-100))


def test_etf_de_renda_fixa_fica_fora_da_apuracao() -> None:
    """Lei 13.043/2014 arts. 2º a 4º: ETF de renda fixa tem tributação exclusiva na fonte; o
    ganho vai para a lista da DIRPF, não para o DARF nem para a isenção."""
    livro = Livro()
    livro.nota(date(2024, 1, 2), livro.op("C", "B5P211", 100, "100"))
    livro.nota(date(2024, 2, 1), livro.op("V", "B5P211", 100, "110"))
    ap = livro.apurar(date(2024, 2, 29))
    assert categoria(ap, 2024, 2, Categoria.COMUM).resultado == 0
    [exclusiva] = ap.exclusiva
    assert (exclusiva.ticker, exclusiva.resultado) == ("B5P211", D(1000))


# ------------------------------------------------------------------- robustez (revisão da F4)
@pytest.mark.parametrize(("preco_venda", "imposto"), [("25", "300.00"), ("15", "0.00")])
def test_venda_a_descoberto_resultado_na_recompra_e_limite_do_mes_da_venda(
    preco_venda: str, imposto: str
) -> None:
    """Venda a descoberto (sem norma específica; leitura A1 escolhida pelo usuário em 15/10/2026:
    resultado na recompra, por analogia com IN 1.585 art. 62 II e P&R 725; limite pelo mês da
    venda). Maio: venda a descoberto de 1.000 BBAS3 (25.000 ou 15.000). Junho: recompra a 23 e
    venda de 200 VALE3 por 5.000 com ganho 200. Com 25.000: os 2.000 são tributáveis (maio acima
    do limite) e VALE3 isenta -> IR 300,00. Com 15.000: recompra a 13 (ganho 2.000) isenta."""
    livro = Livro()
    livro.nota(date(2025, 4, 1), livro.op("C", "VALE3", 200, "24"))
    livro.nota(date(2025, 5, 15), livro.op("V", "BBAS3", 1000, preco_venda))
    recompra = str(Decimal(preco_venda) - 2)
    livro.nota(date(2025, 6, 10), livro.op("C", "BBAS3", 1000, recompra),
               livro.op("V", "VALE3", 200, "25"))  # fmt: skip
    ap = livro.apurar(date(2025, 6, 30))
    assert categoria(ap, 2025, 5, Categoria.COMUM).resultado == 0
    junho = categoria(ap, 2025, 6, Categoria.COMUM)
    assert junho.imposto == D(imposto)
    assert any("descoberto" in a for a in ap.avisos)


def test_evento_que_inverte_a_posicao_falha_alto() -> None:
    livro = Livro()
    livro.nota(date(2024, 1, 2), livro.op("C", "AAAA3", 100, "10"))
    livro.evento(date(2024, 2, 1), "AAAA3", TipoEvento.GRUPAMENTO, "V", "110")
    with pytest.raises(MotorError, match="inverte"):
        livro.movimentos(date(2024, 2, 29))


def test_opcao_com_vencimento_anterior_ao_negocio_falha_alto() -> None:
    livro = Livro()
    livro.nota(date(2024, 3, 1), livro.op("C", "AAAAB50", 100, "1", mercado="CALL",
                                          vencimento=date(2024, 2, 16)))  # fmt: skip
    with pytest.raises(MotorError, match="anterior ao negócio"):
        livro.movimentos(date(2024, 3, 31))


def test_mesmo_codigo_de_serie_em_anos_diferentes_sao_posicoes_separadas() -> None:
    """IN RFB 1.585/2015 art. 58 (custo médio por ativo) e regulamento de opções da B3 (o código
    da série se repete entre anos): TAEEL382 venceu em dez/2023 e em dez/2025; a de 2023 vira pó
    (R10), a de 2025 continua."""
    livro = Livro()
    livro.nota(date(2023, 10, 2), livro.op("V", "TAEEL382", 100, "1", mercado="CALL",
                                           vencimento=date(2023, 12, 15)))  # fmt: skip
    livro.nota(date(2025, 10, 1), livro.op("V", "TAEEL382", 200, "2", mercado="CALL",
                                           vencimento=date(2025, 12, 19)))  # fmt: skip
    ap = livro.apurar(date(2025, 11, 30))
    assert categoria(ap, 2023, 12, Categoria.COMUM).por_grupo[Grupo.OPCOES] == D(100)
    assert categoria(ap, 2025, 10, Categoria.COMUM).resultado == 0
    assert [m.quantidade_apos for m in ap.movimentos][-1] == -200


def test_exercicio_parcial_e_sobra_vira_po_no_vencimento() -> None:
    """R9/R10 (IN RFB 1.585/2015 art. 58; P&R IRPF, opções): titular de 1.000 calls exerce 600 no
    vencimento; as 400 restantes viram pó (perde 40% do prêmio de 1.000 = -400); o custo das
    ações leva 600 do prêmio."""
    livro = Livro()
    livro.nota(date(2024, 1, 10), livro.op("C", "AAAAX30", 1000, "1", mercado="CALL",
                                           vencimento=date(2024, 2, 16)))  # fmt: skip
    livro.nota(date(2024, 2, 16), livro.op("C", "AAAA3", 600, "30", mercado="EXERC CALL",
                                           opcao="AAAAX30"))  # fmt: skip
    ap = livro.apurar(date(2024, 2, 29))
    assert categoria(ap, 2024, 2, Categoria.COMUM).por_grupo[Grupo.OPCOES] == D(-400)
    acao = [m for m in ap.movimentos if m.ticker == "AAAA3"][-1]
    assert acao.custo_apos == D(18600)


def test_bonificacao_ajusta_a_quantidade_da_serie_sem_custo() -> None:
    """R8 (Lei 9.249/1995 art. 10 parágrafo único e IN RFB 1.585/2015 art. 58: o custo atribuído
    da bonificação vai para as ações) e ajuste de proventos da B3 nas séries de opção: a série só
    muda a quantidade, sem custo (1.000 -> 1.050 lançadas)."""
    livro = Livro()
    livro.nota(date(2023, 10, 30), livro.op("V", "ITSAL921", 1000, "0.10", mercado="CALL",
                                            vencimento=date(2023, 12, 15)))  # fmt: skip
    livro.nota(date(2023, 10, 31), livro.op("C", "ITSA4", 1050, "9"))
    livro.evento(date(2023, 11, 29), "ITSAL921", TipoEvento.BONIFICACAO, "V", "50",
                 tipo_ativo=TipoAtivo.OPCAO)  # fmt: skip
    livro.nota(date(2023, 12, 15), livro.op("V", "ITSA4", 1050, "9.21", mercado="EXERC CALL",
                                            opcao="ITSAL921"))  # fmt: skip
    venda = [m for m in livro.movimentos(date(2023, 12, 31)) if m.ticker == "ITSA4"][-1]
    assert venda.valor == D("9670.50") + D(100)  # 1.050 x 9,21 + prêmio recebido


def test_irrf_de_day_trade_e_consumido_primeiro() -> None:
    """IRRF em saldo único (Lei 9.959/2000 art. 8º §4º; IN 1.585 art. 65 §8º; P&R 714): abate o
    imposto total do mês. Consome primeiro o de day trade (não vai para a declaração). Nov:
    IRRF 0,005% 3 e day trade 2, imposto 4 -> sobra 1 de 0,005% e 0 de day trade."""
    livro = Livro()
    livro.nota(date(2023, 11, 1), livro.op("C", "TAEE11", 100, "10"), irrf="3", irrf_dt="2")
    livro.nota(date(2023, 11, 10), livro.op("V", "TAEE11", 100, "10.2667"))  # ganho 26,67: IR 4
    ap = livro.apurar(date(2023, 12, 31))
    assert mes(ap, 2023, 11).irrf_deduzido == D("4.00")
    [anual] = ap.irrf_anual
    assert (anual.saldo_comum, anual.saldo_day_trade) == (D(1), 0)


# ------------------------------------------------------------------------- isenção de PME
PME_ATE = date(2023, 12, 31)


def test_pme_ganho_isento_e_fora_do_limite() -> None:
    """Lei 13.043/2014 art. 16; IN 1.585 arts. 66 e 67 §3º: ganho na venda de ações de PME é
    isento e o valor não entra no limite de R$ 20 mil. Jul/2019: AGRO3 vendida por 30.000 (ganho
    3.000) e ITSA4 por 15.000 (ganho 500): limite 15.000 -> as duas isentas, IR 0."""
    livro = Livro()
    livro.nota(date(2019, 6, 17), livro.op("C", "AGRO3", 1000, "27", pme_ate=PME_ATE),
               livro.op("C", "ITSA4", 1000, "14.5"))  # fmt: skip
    livro.nota(date(2019, 7, 1), livro.op("V", "AGRO3", 1000, "30"),
               livro.op("V", "ITSA4", 1000, "15"))  # fmt: skip
    ap = livro.apurar(date(2019, 7, 31))
    c = categoria(ap, 2019, 7, Categoria.COMUM)
    assert (c.vendas_acoes, c.isento, c.imposto) == (D(15000), D(500), 0)
    assert mes(ap, 2019, 7).isento_pme == D(3000)


def test_pme_perda_nao_compensa_e_venda_depois_de_2023_tributa() -> None:
    """IN 1.585 art. 67 §2º: perda em ação de PME não compensa; art. 66: isenção só até
    31/12/2023."""
    livro = Livro()
    livro.nota(date(2020, 1, 2), livro.op("C", "PRIO3", 200, "30", pme_ate=PME_ATE))
    livro.nota(date(2020, 3, 2), livro.op("V", "PRIO3", 100, "20"))  # perda 1.000
    livro.nota(date(2024, 2, 1), livro.op("V", "PRIO3", 100, "40"))  # ganho 1.000 em 2024
    ap = livro.apurar(date(2024, 2, 29))
    assert categoria(ap, 2020, 3, Categoria.COMUM).prejuizo_a_compensar == 0
    assert mes(ap, 2020, 3).isento_pme == D(-1000)
    assert categoria(ap, 2024, 2, Categoria.COMUM).vendas_acoes == D(4000)


def test_pme_compra_depois_de_2023_junta_ao_lote_antigo() -> None:
    """IN 1.585 art. 66: a isenção só vale para vendas até 31/12/2023. Depois disso o lote
    antigo (comprado no período) e uma compra nova são tributados igualmente: o PM junta os dois
    (500 a 30 + 500 a 50 = PM 40) e a venda de 2024 (45.000 > 20.000) dá 5.000 tributáveis."""
    livro = Livro()
    livro.nota(date(2020, 1, 2), livro.op("C", "PRIO3", 500, "30", pme_ate=PME_ATE))
    livro.nota(date(2024, 2, 1), livro.op("C", "PRIO3", 500, "50"))
    livro.nota(date(2024, 3, 1), livro.op("V", "PRIO3", 1000, "45"))
    ap = livro.apurar(date(2024, 3, 31))
    assert categoria(ap, 2024, 3, Categoria.COMUM).resultado == D(5000)
    assert mes(ap, 2024, 3).isento_pme == 0


def test_pme_day_trade_isento_e_lote_misturado_falha() -> None:
    """Day trade de PME no período: isento (art. 59 §2º I só afasta o inciso I; decisão do
    usuário). Lote isento + lote tributado juntos: controle separado (art. 69-A §4º) não
    implementado, falha alto."""
    livro = Livro()
    livro.nota(date(2020, 6, 16), livro.op("C", "PRIO3", 100, "20", pme_ate=PME_ATE),
               livro.op("V", "PRIO3", 100, "21"), irrf_dt="1")  # fmt: skip
    ap = livro.apurar(date(2020, 6, 30))
    assert categoria(ap, 2020, 6, Categoria.DAY_TRADE).resultado == 0
    assert mes(ap, 2020, 6).isento_pme == D(100)
    livro.nota(date(2021, 1, 4), livro.op("C", "PRIO3", 100, "30", pme_ate=PME_ATE))
    livro.nota(date(2021, 4, 1), livro.op("C", "PRIO3", 100, "50"))  # depois da oferta
    with pytest.raises(MotorError, match="mistura lote"):
        livro.movimentos(date(2021, 4, 30))


def test_evento_sem_posicao_falha_alto_e_subscricao_abre_posicao() -> None:
    """R8 e regra 5 (falhar alto): desdobramento, bonificação ou redução de capital sem posição
    criariam quantidade do nada; só a subscrição (C) abre posição (1.000 a 10,00 = PM 10)."""
    livro = Livro()
    livro.evento(date(2024, 2, 1), "AAAA3", TipoEvento.SUBSCRICAO, "C", "1000", "10000")
    [mov] = livro.movimentos(date(2024, 2, 29))
    assert (mov.quantidade_apos, mov.pm_apos) == (D(1000), D(10))
    livro.evento(date(2024, 3, 1), "BBBB3", TipoEvento.BONIFICACAO, "C", "100", "500")
    with pytest.raises(MotorError, match="sem posição"):
        livro.movimentos(date(2024, 3, 31))


def test_venda_a_termo_falha_alto_com_data_e_ticker() -> None:
    """R18: só a compra a termo está implementada (regra 5: falhar alto, com a nota localizável)."""
    livro = Livro()
    livro.nota(date(2024, 1, 3), livro.op("V", "AAAA3", 100, "31", mercado="TERMO",
                                          liquidacao=date(2024, 3, 4)))  # fmt: skip
    with pytest.raises(MotorError, match="2024-01-03, AAAA3: venda a termo"):
        livro.movimentos(date(2024, 3, 31))
