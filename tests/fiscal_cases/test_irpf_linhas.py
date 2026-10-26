"""Linha da ficha "Renda variável - operações comuns/day trade" do programa IRPF (F8).

Casos do especialista fiscal (pareceres de 20 e 21/10/2026), conferidos à mão.

Fontes: programa IRPF, ficha de renda variável (linhas "Mercado de opções - ações/outros",
"Mercado futuro - juros" e "Mercado a termo - ações/ouro"); P&R IRPF 2026 pergunta 707 (cotas de
fundo de índice não são ações); IN RFB 1.585/2015 art. 62 I e §1º (ganho do comprador a termo pela
venda à vista na data da liquidação; vendido depois, custo = preço a termo) e art. 59 §2º IV (esse
ganho não é isento); art. 58 (custo médio); decisões do usuário de 20/10/2026 (linha do
termo) e de 21/10/2026 (lote do termo vendido primeiro, pela ordem dos negócios).
"""

from datetime import date

from coir.domain.fiscal import Categoria, Grupo, TipoAtivo
from tests.fiscal_cases.apoio import D, Livro, categoria, mes

VENC = date(2025, 4, 17)


def test_opcao_sobre_etf_vai_para_opcoes_outros() -> None:
    """Caso F, março/2025: PETRC350 (objeto PETR4) C 1.000 a 1,00 e V a 1,30 (+300); opção sobre
    BOVA11 C 100 a 3,00 e V a 5,00 (+200). "Opções - ações" 300,00, "opções - outros" 200,00;
    resultado 500,00, imposto 15% = 75,00; IRRF de 0,09 dispensado; DARF vence em 30/04/2025."""
    livro = Livro()
    acao, etf = TipoAtivo.ACAO, TipoAtivo.ETF
    livro.nota(date(2025, 3, 5), livro.op("C", "PETRC350", 1000, "1.00", mercado="CALL",
                                          vencimento=VENC, objeto=acao),
               livro.op("C", "BOVAC130", 100, "3.00", mercado="CALL", vencimento=VENC,
                        objeto=etf))  # fmt: skip
    livro.nota(date(2025, 3, 20), livro.op("V", "PETRC350", 1000, "1.30", mercado="CALL",
                                           vencimento=VENC, objeto=acao),
               livro.op("V", "BOVAC130", 100, "5.00", mercado="CALL", vencimento=VENC,
                        objeto=etf))  # fmt: skip
    ap = livro.apurar(date(2025, 3, 31))
    comum = categoria(ap, 2025, 3, Categoria.COMUM)
    assert (comum.por_grupo[Grupo.OPCOES], comum.por_grupo[Grupo.OPCOES_OUTROS]) == (
        D("300.00"), D("200.00"))  # fmt: skip
    assert (comum.resultado, comum.imposto) == (D("500.00"), D("75.00"))
    m = mes(ap, 2025, 3)
    assert (m.darf, m.vencimento) == (D("75.00"), date(2025, 4, 30))


def test_opcao_sobre_unit_e_sem_objeto_ficam_em_acoes() -> None:
    """Unit acompanha a ação (parâmetro opcoes_linha_acoes_tipos_objeto = ACAO;UNIT, como o
    Bens e Direitos 03.01 da DIRPF 2026); objeto não identificado fica em ações, como antes."""
    livro = Livro()
    livro.nota(date(2025, 3, 5), livro.op("C", "TAEEC360", 100, "1.00", mercado="CALL",
                                          vencimento=VENC, objeto=TipoAtivo.UNIT),
               livro.op("C", "XXXXC100", 100, "1.00", mercado="CALL", vencimento=VENC))  # fmt: skip
    livro.nota(date(2025, 3, 6), livro.op("V", "TAEEC360", 100, "2.00", mercado="CALL",
                                          vencimento=VENC, objeto=TipoAtivo.UNIT),
               livro.op("V", "XXXXC100", 100, "3.00", mercado="CALL", vencimento=VENC))  # fmt: skip
    ap = livro.apurar(date(2025, 3, 31))
    comum = categoria(ap, 2025, 3, Categoria.COMUM)
    assert (comum.por_grupo[Grupo.OPCOES], comum.por_grupo[Grupo.OPCOES_OUTROS]) == (
        D("300.00"), D("0"))  # fmt: skip
    # objeto não identificado não fica calado (revisão da F8)
    [aviso] = [a for a in ap.avisos if "não identificou o ativo-objeto" in a]
    assert aviso.startswith("Opção XXXXC100:") and '"Mercado de opções - ações"' in aviso


def _termo(livro: Livro) -> None:
    livro.nota(date(2025, 11, 10), livro.op("C", "AAAA3", 1000, "10.00", mercado="TERMO",
                                            liquidacao=date(2025, 12, 10)))  # fmt: skip


def test_venda_na_data_da_liquidacao_vai_para_a_linha_do_termo() -> None:
    """Caso D2: compra a termo de 1.000 AAAA3 a 10,00 em 10/11/2025, liquidação em 10/12/2025;
    venda à vista no mesmo dia a 11,00 = ganho 1.000,00 na linha "Mercado a termo", nunca isento
    (vendas de 11.000 abaixo de R$ 20 mil): imposto 150,00; DARF vence em 30/01/2026."""
    livro = Livro()
    _termo(livro)
    livro.nota(date(2025, 12, 10), livro.op("V", "AAAA3", 1000, "11.00"))
    ap = livro.apurar(date(2025, 12, 31))
    comum = categoria(ap, 2025, 12, Categoria.COMUM)
    assert (comum.por_grupo[Grupo.TERMO], comum.por_grupo[Grupo.A_VISTA]) == (D("1000.00"), 0)
    assert (comum.isento, comum.imposto) == (0, D("150.00"))
    m = mes(ap, 2025, 12)
    assert (m.darf, m.vencimento) == (D("150.00"), date(2026, 1, 30))


def test_venda_depois_da_liquidacao_e_a_vista_e_pode_ser_isenta() -> None:
    """Caso D1: o mesmo termo, vendido em 15/01/2026 a 12,00: custo = preço a termo (art. 62
    §1º), ganho 2.000,00 à vista, isento (vendas de 12.000 no mês)."""
    livro = Livro()
    _termo(livro)
    livro.nota(date(2026, 1, 15), livro.op("V", "AAAA3", 1000, "12.00"))
    comum = categoria(livro.apurar(date(2026, 1, 31)), 2026, 1, Categoria.COMUM)
    assert (comum.por_grupo[Grupo.TERMO], comum.isento, comum.imposto) == (0, D("2000.00"), 0)


def _antigas_e_termo(livro: Livro, custos: str = "0") -> None:
    """Parecer do especialista de 21/10/2026: 500 AAAA3 a 9,00 em carteira (custo 4.500,00) e
    termo de 1.000 a 10,00 liquidado em 10/12/2025."""
    livro.nota(date(2025, 10, 1), livro.op("C", "AAAA3", 500, "9.00"))
    termo = livro.op("C", "AAAA3", 1000, "10.00", mercado="TERMO", liquidacao=date(2025, 12, 10))
    livro.nota(date(2025, 11, 10), termo, custos=custos)


def test_venda_no_dia_acima_do_termo_sai_das_antigas_pelo_pm() -> None:
    """C1: venda de 1.500 a 11,00 no dia da liquidação. O lote do termo sai primeiro, pelo preço a
    termo (IN 1.585 art. 62 I): 11.000,00 - 10.000,00 = 1.000,00 na linha do termo, tributável. O
    excedente sai das antigas pelo PM: 5.500,00 - 4.500,00 = 1.000,00, isento (vendas de 16.500,00,
    abaixo de R$ 20 mil; Lei 11.033/2004 art. 3º I). Imposto 150,00; posição zerada (decisão do
    usuário de 21/10/2026)."""
    livro = Livro()
    _antigas_e_termo(livro)
    livro.nota(date(2025, 12, 10), livro.op("V", "AAAA3", 1500, "11.00"))
    comum = categoria(livro.apurar(date(2025, 12, 31)), 2025, 12, Categoria.COMUM)
    assert (comum.por_grupo[Grupo.TERMO], comum.isento, comum.imposto) == (
        D("1000.00"), D("1000.00"), D("150.00"))  # fmt: skip
    assert livro.carteira(date(2025, 12, 31)) == []


def test_varias_vendas_no_dia_consomem_o_termo_pela_ordem() -> None:
    """C2 (variante): 1.000 a 11,00 e depois 500 a 12,00 no dia da liquidação. Pela ordem dos
    negócios (decisão do usuário de 21/10/2026), a primeira venda é do termo (1.000,00 na linha do
    termo) e a segunda das antigas: 6.000,00 - 4.500,00 = 1.500,00, isento (vendas de 17.000,00).
    Imposto 150,00."""
    livro = Livro()
    _antigas_e_termo(livro)
    livro.nota(date(2025, 12, 10), livro.op("V", "AAAA3", 1000, "11.00"),
               livro.op("V", "AAAA3", 500, "12.00"))  # fmt: skip
    comum = categoria(livro.apurar(date(2025, 12, 31)), 2025, 12, Categoria.COMUM)
    assert (comum.por_grupo[Grupo.TERMO], comum.isento, comum.imposto) == (
        D("1000.00"), D("1500.00"), D("150.00"))  # fmt: skip


def test_venda_no_dia_menor_que_o_termo_usa_o_preco_a_termo() -> None:
    """C3: venda de 600 a 11,00 no dia da liquidação, com custos de 30,00 na nota do termo. As 600
    saem pelo preço a termo + custos (art. 62 I): 6.600,00 - 6.018,00 = 582,00 na linha do termo,
    imposto 87,30. As 400 que sobram entram no custo médio pelo preço a termo (art. 62 §1º e art.
    58): 900 ações com custo 4.500,00 + 4.012,00 = 8.512,00. Pelo PM misturado o ganho seria
    maior e o custo também (achado do parecer de 21/10/2026)."""
    livro = Livro()
    _antigas_e_termo(livro, custos="30")
    livro.nota(date(2025, 12, 10), livro.op("V", "AAAA3", 600, "11.00"))
    ap = livro.apurar(date(2025, 12, 31))
    comum = categoria(ap, 2025, 12, Categoria.COMUM)
    assert (comum.por_grupo[Grupo.TERMO], comum.isento, comum.imposto) == (
        D("582.00"), 0, D("87.30"))  # fmt: skip
    m = mes(ap, 2025, 12)
    assert (m.darf, m.vencimento) == (D("87.30"), date(2026, 1, 30))
    [pos] = livro.carteira(date(2025, 12, 31))
    assert (pos.quantidade, pos.custo) == (900, D("8512.00"))


def test_venda_do_termo_entra_na_soma_do_limite_de_isencao() -> None:
    """Decisão de 15/10/2026 (F4), mantida na F8: o ganho da venda no dia da liquidação do termo
    nunca é isento (IN 1.585 art. 59 §2º IV), mas o valor da venda entra na soma das alienações
    do mês (leitura conservadora; nenhuma fonte decide). Dezembro: venda do termo de 11.000,00 e
    venda de 1.000 BBBB3 a 10,00 = 10.000,00 (PM 9,00: ganho 1.000,00); vendas de 21.000,00 >
    20.000,00: o ganho de BBBB3 também é tributável. Imposto 15% x (1.000 + 1.000) = 300,00."""
    livro = Livro()
    livro.nota(date(2025, 10, 1), livro.op("C", "BBBB3", 1000, "9.00"))
    _termo(livro)
    livro.nota(date(2025, 12, 10), livro.op("V", "AAAA3", 1000, "11.00"),
               livro.op("V", "BBBB3", 1000, "10.00"))  # fmt: skip
    comum = categoria(livro.apurar(date(2025, 12, 31)), 2025, 12, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.isento, comum.imposto) == (D("21000"), 0, D("300.00"))
