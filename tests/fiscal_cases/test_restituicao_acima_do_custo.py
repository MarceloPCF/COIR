"""Restituição de capital (redução de capital com devolução em dinheiro) e amortização de cotas
maiores que o custo da posição (parecer do especialista fiscal de 24/11/2026; números conferidos à
mão).

- Ação e unit: a restituição reduz o custo (IN RFB 1.585/2015 art. 58 §8º). O que passar do custo
  não vira custo negativo (Lei 7.713/1988 art. 16; RIR/2018 art. 141 §2º): custo e PM a zero, e o
  excedente é ganho de capital fora de bolsa (Lei 7.713/1988 art. 3º §§2º a 4º; SC Cosit 37/2020 por
  analogia; SC Cosit 131/2016, ementa), no GCAP, no mês do recebimento, fora da apuração mensal e do
  limite de R$ 20 mil (IN 1.585 art. 59 I); não compensa prejuízo de bolsa (SC Cosit 130/2023).
  Critério do COIR, sem norma expressa (parâmetro restituicao_capital_acima_do_custo; alternativa:
  isolar).
- Venda posterior: ganho = valor da venda - custo zero (IN 1.585 art. 58 caput).
- Custo médio único do titular (IN 1.585 art. 58 caput; P&R IRPF 2026 pergunta 720): o excedente se
  mede contra o custo somado de todas as contas.
- FII e Fiagro: sem norma expressa para a amortização; por analogia (IN 1.585 arts. 16 §2º e 37 §3º)
  reduz o custo, e o excedente é avisado (parâmetro fii_amortizacao).
- ETF de ações desde 2024: só a parcela custo x amortização / valor patrimonial devolve custo (Lei
  14.754/2023 arts. 18 II, 24 e 17 §2º III e §5º II b).
- BDR acima do custo: sem norma; o ativo fica isolado.
"""

from datetime import date

from coir.domain.fiscal import Categoria, TipoAtivo, TipoEvento, Visao
from tests.fiscal_cases.apoio import D, Livro, categoria, parametros

REST = TipoEvento.REDUCAO_CAPITAL
EX = date(2025, 7, 1)  # data ex; pagamento em 15/07/2025


def _acao(quantidade: int, custos: str = "0.30", valor: str = "1500") -> Livro:
    """10/03/2025: C XPTO3 a 1,20 (custos da nota) e C 1.000 ABCD3 a 15,00. Restituição de R$ 1,50
    por ação na data ex 01/07/2025."""
    livro = Livro()
    livro.nota(date(2025, 3, 10), livro.op("C", "XPTO3", quantidade, "1.20"), custos=custos)
    livro.nota(date(2025, 3, 10), livro.op("C", "ABCD3", 1000, "15"))
    livro.evento(EX, "XPTO3", REST, "V", None, valor)
    return livro


def _custo(livro: Livro, ticker: str, dia: date) -> tuple[D, D]:
    [p] = [p for p in livro.carteira(dia) if p.ticker == ticker]
    return p.quantidade, p.custo


def test_r1_excedente_zera_o_custo_e_fica_fora_da_apuracao() -> None:
    """Custo 1.200,30; restituição 1.500,00; excedente 299,70. Julho: V 1.000 ABCD3 a 19,00 =
    19.000,00 (ganho 4.000,00) segue isenta; somada a restituição, a soma passaria de R$ 20 mil e o
    imposto seria 600,00. GCAP: 1.500,00 recebidos no mês, até R$ 35 mil (Lei 9.250/1995 art. 22
    II): isento, como no resgate PETZ/AUAU."""
    livro = _acao(1000)
    livro.nota(date(2025, 7, 15), livro.op("V", "ABCD3", 1000, "19"))
    assert _custo(livro, "XPTO3", date(2025, 7, 31)) == (D(1000), D(0))
    ap = livro.apurar(date(2025, 7, 31))
    julho = categoria(ap, 2025, 7, Categoria.COMUM)
    assert (julho.vendas_limite, julho.isento, julho.imposto) == (D(19000), D(4000), D(0))
    assert not ap.isolados
    [aviso] = [a for a in ap.avisos if "acima do custo" in a]
    assert aviso.startswith("01/07/2025: restituição de capital de XPTO3 acima do custo")
    assert "R$ 299,70" in aviso and "GCAP" in aviso


def test_r2_excedente_grande_e_venda_posterior_tributada() -> None:
    """100.000 a 1,20 = 120.000,00; restituição 150.000,00: excedente 30.000,00 -> GCAP 15% =
    4.500,00, DARF 4600 até 29/08/2025 (fora do COIR). 10/09: V 100.000 a 2,00 = 200.000,00 com
    custo zero: ganho 200.000,00, imposto 30.000,00 (DARF 6015 até 31/10/2025)."""
    livro = _acao(100000, custos="0", valor="150000")
    assert D("150000") - D("120000") == D(30000)
    assert (D(30000) * D("0.15")).quantize(D("0.01")) == D("4500.00")
    livro.nota(date(2025, 9, 10), livro.op("V", "XPTO3", 100000, "2"))
    setembro = categoria(livro.apurar(date(2025, 9, 30)), 2025, 9, Categoria.COMUM)
    assert (setembro.resultado, setembro.isento, setembro.imposto) == (
        D(200000), D(0), D("30000.00"))  # fmt: skip


def test_r3_venda_pequena_com_custo_zero_e_isenta() -> None:
    """10/09/2025: V 1.000 XPTO3 a 2,00 = 2.000,00; custo zero: ganho 2.000,00, isento."""
    livro = _acao(1000)
    livro.nota(date(2025, 9, 10), livro.op("V", "XPTO3", 1000, "2"))
    setembro = categoria(livro.apurar(date(2025, 9, 30)), 2025, 9, Categoria.COMUM)
    assert (setembro.resultado, setembro.isento, setembro.imposto) == (D(0), D(2000), D(0))


def test_r4_excedente_se_mede_contra_o_custo_do_titular() -> None:
    """XP: C 600 a 1,20 = 720,00; Clear: C 400 a 2,00 = 800,00; titular: custo 1.520,00.
    Restituição de 900,00 (XP) e 600,00 (Clear): sobra custo de 20,00 no titular, sem aviso. Na
    visão por conta a XP passaria do custo: a visão operacional não falha nem avisa (o custo fiscal
    é o do titular)."""
    livro = Livro()
    livro.nota(date(2025, 3, 10), livro.op("C", "XPTO3", 600, "1.20"))
    livro.nota(date(2025, 3, 11), livro.op("C", "XPTO3", 400, "2"), corretora="CLEAR", conta=2)
    livro.evento(EX, "XPTO3", REST, "V", None, "900")
    livro.evento(EX, "XPTO3", REST, "V", None, "600", conta=2, corretora="CLEAR")
    assert _custo(livro, "XPTO3", date(2025, 7, 31)) == (D(1000), D(20))
    ap = livro.apurar(date(2025, 7, 31))
    assert not ap.isolados and not [a for a in ap.avisos if "acima do custo" in a]
    por_conta = {p.conta: p.custo for p in livro.carteira(date(2025, 7, 31), Visao.CONTA)}
    assert por_conta == {1: D(0), 2: D(200)}


def test_r5_amortizacao_de_fii_reduz_o_custo() -> None:
    """C 100 FIIX11 a 100,00 = 10.000,00; amortização de 200,00 em 15/05: custo 9.800,00. V 100 a
    99,00 = 9.900,00: ganho 100,00 x 20% = 20,00. Sem a amortização: prejuízo de 100,00."""
    fii = TipoAtivo.FII
    livro = Livro()
    livro.nota(date(2025, 2, 5), livro.op("C", "FIIX11", 100, "100", tipo=fii))
    livro.evento(date(2025, 5, 15), "FIIX11", REST, "V", None, "200", tipo_ativo=fii)
    assert _custo(livro, "FIIX11", date(2025, 5, 31)) == (D(100), D(9800))
    livro.nota(date(2025, 6, 10), livro.op("V", "FIIX11", 100, "99", tipo=fii))
    junho = categoria(livro.apurar(date(2025, 6, 30)), 2025, 6, Categoria.FII)
    assert (junho.resultado, junho.imposto) == (D(100), D("20.00"))


def test_r6_amortizacao_de_fii_acima_do_custo_zera_e_avisa() -> None:
    """C 10 FIIY11 a 100,00 = 1.000,00; amortização de 1.050,00: custo zero, excedente 50,00."""
    fii = TipoAtivo.FII
    livro = Livro()
    livro.nota(date(2025, 2, 5), livro.op("C", "FIIY11", 10, "100", tipo=fii))
    livro.evento(date(2025, 5, 15), "FIIY11", REST, "V", None, "1050", tipo_ativo=fii)
    assert _custo(livro, "FIIY11", date(2025, 5, 31)) == (D(10), D(0))
    ap = livro.apurar(date(2025, 5, 31))
    assert not ap.isolados
    [aviso] = [a for a in ap.avisos if "acima do custo" in a]
    assert "R$ 50,00" in aviso and "20%" in aviso


def test_r7_amortizacao_de_etf_desde_2024_pela_parcela_de_custo() -> None:
    """C 100 ETFX11 a 100,00; amortização de 500,00 com VP 125,00: parcela de custo 400,00 (Lei
    14.754/2023 art. 17 §5º II b); custo 9.600,00. V 100 a 101,00: ganho 500,00, imposto 75,00.
    Com o valor inteiro: custo 9.500,00, ganho 600,00, imposto 90,00 e aviso."""
    etf = TipoAtivo.ETF
    for valor, ganho, imposto in (("400", "500", "75.00"), ("500", "600", "90.00")):
        livro = Livro()
        livro.nota(date(2025, 2, 5), livro.op("C", "ETFX11", 100, "100", tipo=etf))
        livro.evento(date(2025, 6, 16), "ETFX11", REST, "V", None, valor, tipo_ativo=etf)
        livro.nota(date(2025, 9, 10), livro.op("V", "ETFX11", 100, "101", tipo=etf))
        ap = livro.apurar(date(2025, 9, 30))
        setembro = categoria(ap, 2025, 9, Categoria.COMUM)
        assert (setembro.resultado, setembro.imposto) == (D(ganho), D(imposto))
        assert any("parcela de custo" in a for a in ap.avisos)


def test_r8_bdr_acima_do_custo_fica_isolado() -> None:
    bdr = TipoAtivo.BDR
    livro = Livro()
    livro.nota(date(2025, 3, 10), livro.op("C", "XXXX34", 10, "10", tipo=bdr))
    livro.evento(EX, "XXXX34", REST, "V", None, "120", tipo_ativo=bdr)
    [isolado] = livro.apurar(date(2025, 7, 31)).isolados
    assert isolado.ticker == "XXXX34" and isolado.desde == EX
    assert "maior que o custo" in isolado.motivo


def test_r9_com_o_parametro_isolar_a_acao_sai_do_calculo() -> None:
    ap = _acao(1000).apurar(date(2025, 7, 31),
                            parametros(restituicao_capital_acima_do_custo="isolar"))  # fmt: skip
    assert [i.ticker for i in ap.isolados] == ["XPTO3"]
