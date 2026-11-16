"""Cotas de FIP comum e de FIDC vendidas em bolsa por pessoa física (item B6 do parecer do
especialista de 13/11/2026).

- Cota de fundo não é ação: fora da soma do limite de R$ 20 mil e sem isenção (Lei 11.033/2004
  art. 3º I; SC Cosit 145/2021 e 166/2021).
- FIP comum: ganho líquido a 15% na venda em bolsa (IN RFB 1.585/2015 art. 32 §1º I a e §3º; Lei
  11.312/2006 art. 2º §1º I, não revogado pela Lei 14.754/2023 art. 46), na categoria comum (Lei
  8.981/1995 art. 72 §4º).
- Amortização de FIP até 31/12/2023: reduz o custo; só o excedente é rendimento (IN 1.585 art. 32
  §2º). Desde 01/01/2024: só a parcela custo x amortização / valor patrimonial devolve custo; o
  restante é rendimento com IRRF de 15%, definitivo (Lei 14.754/2023 art. 24 caput e §2º c/c art.
  17 §2º III e §5º II b; art. 32 I). O COIR não tem o valor patrimonial: o usuário lança a
  parcela de custo do informe; com o valor inteiro, o custo cai até zero, com aviso.
- FIDC: sem norma expressa para a venda em bolsa por pessoa física (o art. 16 I a da IN 1.585 só
  alcança fundo fechado com carteira de ações, art. 18 §2º): o COIR isola o ativo, com o motivo
  (parâmetro fidc_venda_bolsa; ganho_liquido apura a 15%, como a prática de mercado).

Casos T6-1 a T6-3 do parecer, conferidos à mão; sem IRRF, para simplificar."""

from datetime import date

from coir.domain.fiscal import Categoria, TipoAtivo, TipoEvento
from tests.fiscal_cases.apoio import D, Livro, categoria, parametros

FIP, FIDC = TipoAtivo.FIP, TipoAtivo.FIDC
AMORTIZACAO = TipoEvento.REDUCAO_CAPITAL


def _fundo_e_acoes(tipo: TipoAtivo) -> Livro:
    """05/02/2026: C 100 FINC11 a 95,00 + 0,50 = 9.500,50. 10/03: V 100 a 101,00 - 0,50 =
    10.099,50. No mesmo mês: V 1.000 ITSA4 a 15,00 (15.000,00), compradas a 13,00: ganho de
    2.000,00."""
    livro = Livro()
    livro.nota(date(2026, 1, 5), livro.op("C", "ITSA4", 1000, "13"))
    livro.nota(date(2026, 2, 5), livro.op("C", "FINC11", 100, "95", tipo=tipo), custos="0.50")
    livro.nota(date(2026, 3, 10), livro.op("V", "FINC11", 100, "101", tipo=tipo), custos="0.50")
    livro.nota(date(2026, 3, 12), livro.op("V", "ITSA4", 1000, "15"))
    return livro


def test_t6_1_fip_paga_15_por_cento_e_nao_entra_no_limite_das_acoes() -> None:
    """Ganho do FIP: 10.099,50 - 9.500,50 = 599,00, tributado (sem isenção). Soma do limite: só
    os 15.000,00 de ITSA4, abaixo de 20.000,00: os 2.000,00 das ações são isentos. Imposto: 15%
    de 599,00 = 89,85. Como ação, a soma seria 25.100,00 e as ações pagariam 300,00 a mais."""
    ap = _fundo_e_acoes(FIP).apurar(date(2026, 3, 31))
    comum = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (comum.resultado, comum.isento, comum.imposto) == (D(599), D(2000), D("89.85"))
    assert comum.vendas_limite == D(15000)
    assert not ap.isolados
    assert all(not m.isentavel and not m.venda_acao_a_vista
               for m in ap.movimentos if m.ticker == "FINC11")  # fmt: skip


def test_t6_3_fidc_fica_isolado_com_o_motivo_e_fora_do_limite() -> None:
    """Sem norma expressa: o resultado do FIDC não é calculado. As ações seguem isentas (soma do
    limite 15.000,00) e não há imposto no mês."""
    ap = _fundo_e_acoes(FIDC).apurar(date(2026, 3, 31))
    [isolado] = ap.isolados
    assert isolado.ticker == "FINC11" and isolado.desde == date(2026, 2, 5)
    assert "não há norma expressa para a venda de FIDC" in isolado.motivo
    comum = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (comum.isento, comum.imposto) == (D(2000), D(0))
    assert comum.vendas_limite == D(15000)


def test_fidc_com_o_parametro_da_pratica_de_mercado_e_apurado_a_15_por_cento() -> None:
    ap = _fundo_e_acoes(FIDC).apurar(date(2026, 3, 31),
                                     parametros(fidc_venda_bolsa="ganho_liquido"))  # fmt: skip
    comum = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (comum.resultado, comum.isento, comum.imposto) == (D(599), D(2000), D("89.85"))
    assert not ap.isolados


def test_fip_antes_da_vigencia_do_parametro_fica_isolado_sem_palpite() -> None:
    """Sem regra vigente na data (antes da IN RFB 1.585/2015), o COIR não apura."""
    livro = Livro()
    livro.nota(date(2014, 2, 5), livro.op("C", "FINC11", 100, "95", tipo=FIP))
    [isolado] = livro.apurar(date(2014, 12, 31)).isolados
    assert "não tem regra" in isolado.motivo


# ------------------------------------------------------------------ T6-2: amortização de FIP
def _com_amortizacao(compra: date, amortizacao: date, valor: str) -> Livro:
    livro = Livro()
    livro.nota(compra, livro.op("C", "FINC11", 100, "95", tipo=FIP), custos="0.50")
    livro.evento(amortizacao, "FINC11", AMORTIZACAO, "V", valor=valor, tipo_ativo=FIP)
    return livro


def _custo(livro: Livro, dia: date) -> D:
    [p] = [p for p in livro.carteira(dia) if p.ticker == "FINC11"]
    return p.custo


def test_t6_2_amortizacao_ate_2023_reduz_o_custo_sem_rendimento() -> None:
    """Custo 9.500,50; amortização de 12,00 x 100 = 1.200,00 em 2023: custo 8.300,50. Só o que
    exceder o custo seria rendimento (IN RFB 1.585/2015 art. 32 §2º)."""
    livro = _com_amortizacao(date(2023, 2, 6), date(2023, 6, 15), "1200")
    assert _custo(livro, date(2023, 6, 30)) == D("8300.50")
    assert not [a for a in livro.apurar(date(2023, 6, 30)).avisos if "amortização" in a]


def test_t6_2_amortizacao_desde_2024_pela_parcela_de_custo_do_informe() -> None:
    """2025, valor patrimonial da cota 110,00: parcela de custo 9.500,50 x 1.200 / 11.000 =
    1.036,42 (o informe mostra; rendimento 163,58, IRRF 24,54 na fonte). Lançada a parcela, o
    custo fica 8.464,08. Venda por 10.099,50: ganho 1.635,42, imposto 245,31."""
    livro = _com_amortizacao(date(2025, 2, 5), date(2025, 6, 16), "1036.42")
    assert _custo(livro, date(2025, 6, 30)) == D("8464.08")
    livro.nota(date(2025, 9, 10), livro.op("V", "FINC11", 100, "101", tipo=FIP), custos="0.50")
    ap = livro.apurar(date(2025, 9, 30))
    comum = categoria(ap, 2025, 9, Categoria.COMUM)
    assert (comum.resultado, comum.imposto) == (D("1635.42"), D("245.31"))
    assert any("lance só a parcela de custo" in a for a in ap.avisos)
    assert any("come-cotas" in a for a in ap.avisos)  # Lei 14.754/2023 arts. 26 e 17 §2º II


def test_t6_2_amortizacao_desde_2024_pelo_valor_inteiro_e_conservadora() -> None:
    """Sem a parcela do informe, o valor inteiro (1.200,00) reduz o custo: 8.300,50. Venda por
    10.099,50: ganho 1.799,00, imposto 269,85 (a mais, nunca a menos), com aviso."""
    livro = _com_amortizacao(date(2025, 2, 5), date(2025, 6, 16), "1200")
    livro.nota(date(2025, 9, 10), livro.op("V", "FINC11", 100, "101", tipo=FIP), custos="0.50")
    ap = livro.apurar(date(2025, 9, 30))
    comum = categoria(ap, 2025, 9, Categoria.COMUM)
    assert (comum.resultado, comum.imposto) == (D("1799.00"), D("269.85"))
    assert any("lance só a parcela de custo" in a for a in ap.avisos)


def test_amortizacao_de_fip_acima_do_custo_zera_o_custo_e_avisa_o_excedente() -> None:
    """Custo 9.500,50; amortização de 10.000,00: o custo vai a zero e os 499,50 excedentes são
    rendimento tributado na fonte pelo administrador (IN RFB 1.585/2015 art. 32 §2º)."""
    livro = _com_amortizacao(date(2023, 2, 6), date(2023, 6, 15), "10000")
    assert _custo(livro, date(2023, 6, 30)) == D(0)
    [aviso] = [a for a in livro.apurar(date(2023, 6, 30)).avisos if "acima do custo" in a]
    assert "R$ 499,50" in aviso and "tributado na fonte" in aviso
