"""F7a: incorporação, cisão, ajuste de série de opção e leilão de fração (números conferidos à
mão no parecer do especialista fiscal de 19/10/2026).

Fontes:
- IN RFB 1.585/2015 art. 58 §6º: na substituição de ações por incorporação, fusão ou cisão, o
  custo das ações originais é atribuído às recebidas na proporção fixada pela assembleia; P&R
  IRPF (substituição de ações): não é alienação, a data de aquisição é a original.
- Incorporação de ações (Lei 6.404 art. 252): custo mantido (TRF1, TRF3, TRF4; decisão do usuário
  de 19/10/2026); a RFB diverge (SC Cosit 224/2014), parâmetro incorporacao_acoes_tratamento.
- Ajuste de série de opção pela B3 (Manual de Procedimentos Operacionais da Câmara, item 7.9):
  sem efeito fiscal; o custo é o dos prêmios (art. 60 §2º), só a quantidade e o código mudam.
- Fração em leilão: alienação pelo custo médio (art. 58 caput), sem norma sobre a isenção:
  parâmetro leilao_fracao_isentavel (decisão do usuário de 19/10/2026: entra no limite de
  R$ 20 mil do mês do crédito, Lei 11.033/2004 art. 3º I).
"""

from datetime import date
from decimal import Decimal

import pytest

from coir.domain.fiscal import Categoria, NaturezaEvento, PosicaoAtivo, TipoAtivo, TipoEvento, Visao
from coir.erros import MotorError
from tests.fiscal_cases.apoio import Livro, categoria, mes, parametros

D = Decimal
INC = TipoEvento.INCORPORACAO
SOCIEDADE = NaturezaEvento.INCORPORACAO_SOCIEDADE
SERIE = NaturezaEvento.AJUSTE_SERIE
OPCAO = TipoAtivo.OPCAO


Posicoes = dict[tuple[int | None, str], tuple[Decimal, Decimal]]


def _posicoes(carteira: list[PosicaoAtivo]) -> Posicoes:
    return {(p.conta, p.ticker): (p.quantidade, p.custo) for p in carteira}


def _c1(livro: Livro, ja_tinha_b: bool = False) -> None:
    if ja_tinha_b:
        livro.nota(date(2025, 1, 10), livro.op("C", "BBBB3", 10, "38"))  # 10 / 380,00
    livro.nota(date(2025, 2, 3), livro.op("C", "AAAA3", 50, "19.90"), custos="1.00")
    livro.nota(date(2025, 3, 10), livro.op("C", "AAAA3", 25, "20.10"), custos="1.50")
    livro.evento(date(2025, 5, 5), "AAAA3", INC, quantidade="75", para="BBBB3",
                 quantidade_destino="37.5", natureza=SOCIEDADE)  # fmt: skip
    livro.evento(date(2025, 5, 23), "BBBB3", TipoEvento.LEILAO_FRACAO, quantidade="0.5",
                 valor="21.37")  # crédito do leilão (decisão P2)  # fmt: skip


def test_c1_incorporacao_leva_o_custo_e_a_fracao_vendida_entra_no_limite() -> None:
    livro = Livro()
    _c1(livro)
    assert _posicoes(livro.carteira(date(2025, 5, 5))) == {(None, "BBBB3"): (D("37.5"), D(1500))}
    assert _posicoes(livro.carteira(date(2025, 5, 31))) == {(None, "BBBB3"): (D(37), D(1480))}
    movs = livro.movimentos(date(2025, 5, 31))
    assert [m.resultado for m in movs if m.origem.startswith("evento:")] == [0, 0, D("1.37")]
    ap = livro.apurar(date(2025, 5, 31))
    comum = categoria(ap, 2025, 5, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.isento, comum.imposto) == (D("21.37"), D("1.37"), 0)


def test_c1_fracao_nao_isentavel_pelo_parametro() -> None:
    livro = Livro()
    _c1(livro)
    ap = livro.apurar(date(2025, 6, 30), parametros(leilao_fracao_isentavel="false"))
    comum = categoria(ap, 2025, 5, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.resultado, comum.imposto) == (0, D("1.37"), D("0.20"))
    assert mes(ap, 2025, 5).darf == 0  # abaixo de R$ 10: acumula


def test_c1b_b_ja_na_carteira_junta_pela_media_e_a_fracao_sai_pelo_pm() -> None:
    livro = Livro()
    _c1(livro, ja_tinha_b=True)
    [antes] = livro.carteira(date(2025, 5, 5))
    assert (antes.quantidade, antes.custo) == (D("47.5"), D(1880))
    venda = livro.movimentos(date(2025, 5, 31))[-1]
    assert venda.resultado.quantize(D("0.01")) == D("1.58")
    assert venda.custo_apos.quantize(D("0.00000001")) == D("1860.21052632")
    assert venda.quantidade_apos == 47


def test_venda_de_b_antes_da_data_do_evento_falha_alto() -> None:
    livro = Livro()
    livro.nota(date(2025, 2, 3), livro.op("C", "AAAA3", 75, "20"))
    livro.nota(date(2025, 5, 2), livro.op("V", "BBBB3", 37, "40"))  # B vendido antes da data ex
    livro.evento(date(2025, 5, 5), "AAAA3", INC, quantidade="75", para="BBBB3",
                 quantidade_destino="37.5", natureza=SOCIEDADE)  # fmt: skip
    with pytest.raises(MotorError, match="sentido contrário"):
        livro.movimentos(date(2025, 5, 31))


def test_incorporacao_de_acoes_pela_leitura_da_receita_nao_implementada_falha_alto() -> None:
    livro = Livro()
    livro.nota(date(2025, 2, 3), livro.op("C", "AAAA3", 75, "20"))
    livro.evento(
        date(2025, 5, 5),
        "AAAA3",
        INC,
        quantidade="75",
        para="BBBB3",
        quantidade_destino="60.69",
        natureza=NaturezaEvento.INCORPORACAO_ACOES,
    )
    [b] = livro.carteira(date(2025, 5, 31))
    assert (b.ticker, b.custo) == ("BBBB3", D(1500))  # custo mantido (padrão)
    with pytest.raises(MotorError, match="não implementado"):
        livro.apurar(date(2025, 5, 31), parametros(incorporacao_acoes_tratamento="rfb"))


# --------------------------------------------------------------------------------- cisão
def _c2(livro: Livro, x: int = 1, percentual: str = "0.30", so_conta_1: bool = False) -> None:
    livro.nota(date(2025, 6, 2), livro.op("C", "CCCC3", 120 * x, "24.95"), custos=str(6 * x))
    livro.nota(date(2025, 6, 2), livro.op("C", "CCCC3", 80 * x, "24.95"), custos=str(4 * x),
               conta=2)  # fmt: skip
    livro.evento(
        date(2025, 7, 1),
        "CCCC3",
        TipoEvento.CISAO,
        quantidade=str(120 * x),
        para="DDDD3",
        quantidade_destino=str(30 * x),
        percentual=percentual,
    )
    if not so_conta_1:
        livro.evento(date(2025, 7, 1), "CCCC3", TipoEvento.CISAO, quantidade=str(80 * x),
                     para="DDDD3", quantidade_destino=str(20 * x), percentual=percentual,
                     conta=2)  # fmt: skip


def test_c2_cisao_reparte_o_custo_de_antes_do_evento_entre_as_contas() -> None:
    """Visão por titular: 200 / 5.000,00 -> CCCC3 3.500,00 e DDDD3 1.500,00. Aplicada conta a
    conta sobre o custo corrente daria DDDD3 1.392,00 (900 + 4.100 x 0,4 x 0,3): proibido."""
    livro = Livro()
    _c2(livro)
    assert _posicoes(livro.carteira(date(2025, 7, 1))) == {
        (None, "CCCC3"): (D(200), D(3500)), (None, "DDDD3"): (D(50), D(1500))}  # fmt: skip
    assert _posicoes(livro.carteira(date(2025, 7, 1), Visao.CONTA)) == {
        (1, "CCCC3"): (D(120), D(2100)), (1, "DDDD3"): (D(30), D(900)),
        (2, "CCCC3"): (D(80), D(1400)), (2, "DDDD3"): (D(20), D(600))}  # fmt: skip
    livro.nota(date(2025, 7, 15), livro.op("V", "DDDD3", 50, "32"),
               livro.op("V", "CCCC3", 200, "18"), custos="2.60")  # fmt: skip
    vendas = [m for m in livro.movimentos(date(2025, 7, 31)) if m.data == date(2025, 7, 15)]
    assert [(m.ticker, m.resultado) for m in vendas] == [("DDDD3", D("99.20")),
                                                       ("CCCC3", D("98.20"))]  # fmt: skip
    comum = categoria(livro.apurar(date(2025, 7, 31)), 2025, 7, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.isento, comum.imposto) == (D(5200), D("197.40"), 0)


def test_c2b_cisao_com_vendas_acima_do_limite() -> None:
    livro = Livro()
    _c2(livro, x=10)
    livro.nota(date(2025, 7, 15), livro.op("V", "DDDD3", 500, "32"),
               livro.op("V", "CCCC3", 2000, "18"), custos="26", irrf="2.60")  # fmt: skip
    ap = livro.apurar(date(2025, 7, 31))
    comum = categoria(ap, 2025, 7, Categoria.COMUM)
    assert (comum.resultado, comum.imposto) == (D(1974), D("296.10"))
    julho = mes(ap, 2025, 7)
    assert (julho.darf, julho.vencimento) == (D("293.50"), date(2025, 8, 29))


def test_c2c_percentual_com_todas_as_casas_fecha_o_custo() -> None:
    livro = Livro()
    _c2(livro, percentual="0.2987654321")
    pos = _posicoes(livro.carteira(date(2025, 7, 1)))
    assert pos[(None, "DDDD3")] == (D(50), D("1493.8271605"))
    assert pos[(None, "CCCC3")] == (D(200), D("3506.1728395"))
    assert pos[(None, "DDDD3")][1] + pos[(None, "CCCC3")][1] == D(5000)


def test_cisao_de_uma_conta_so_na_visao_por_titular_falha_alto() -> None:
    livro = Livro()
    _c2(livro, so_conta_1=True)
    with pytest.raises(MotorError, match="lance o evento de cada conta"):
        livro.carteira(date(2025, 7, 1))


def test_cisao_sem_percentual_bloqueia_ate_informar() -> None:
    """Decisão do usuário (P4, 19/10/2026): sem o percentual divulgado, não calcula."""
    livro = Livro()
    livro.nota(date(2025, 6, 2), livro.op("C", "CCCC3", 100, "25"))
    livro.evento(date(2025, 7, 1), "CCCC3", TipoEvento.CISAO, quantidade="100", para="DDDD3",
                 quantidade_destino="25")  # fmt: skip
    with pytest.raises(MotorError, match="sem o percentual de custo"):
        livro.carteira(date(2025, 7, 1))


# ---------------------------------------------------------------- ajuste de série da B3
def test_c3a_troca_de_codigo_da_serie_mantem_custo_e_vencimento() -> None:
    livro = Livro()
    venc = date(2025, 1, 17)
    livro.nota(date(2024, 12, 2), livro.op("C", "XXXXA180", 1000, "0.50", mercado="CALL",
                                           vencimento=venc), custos="2.00")  # fmt: skip
    livro.evento(date(2024, 12, 16), "XXXXA180", INC, quantidade="1000", para="XXXXA17",
                 quantidade_destino="1000", natureza=SERIE, tipo_ativo=OPCAO)  # fmt: skip
    [serie] = livro.carteira(date(2024, 12, 16))
    assert (serie.ticker, serie.vencimento, serie.quantidade, serie.custo) == (
        "XXXXA17",
        venc,
        1000,
        D(502),
    )
    livro.nota(date(2025, 1, 10), livro.op("V", "XXXXA17", 1000, "0.80", mercado="CALL",
                                           vencimento=venc), custos="1.50")  # fmt: skip
    ap = livro.apurar(date(2025, 1, 31))
    comum = categoria(ap, 2025, 1, Categoria.COMUM)
    assert (comum.resultado, comum.imposto) == (D("296.50"), D("44.47"))
    janeiro = mes(ap, 2025, 1)
    assert (janeiro.darf, janeiro.vencimento) == (D("44.47"), date(2025, 2, 28))


def test_c3a_lancador_converte_com_o_sinal_e_ganha_no_po() -> None:
    livro = Livro()
    venc = date(2025, 1, 17)
    livro.nota(date(2024, 12, 2), livro.op("V", "XXXXA180", 1000, "0.50", mercado="CALL",
                                           vencimento=venc), custos="2.00")  # fmt: skip
    livro.evento(date(2024, 12, 16), "XXXXA180", INC, quantidade="1000", para="XXXXA17",
                 quantidade_destino="1000", natureza=SERIE, tipo_ativo=OPCAO)  # fmt: skip
    [serie] = livro.carteira(date(2024, 12, 16))
    assert (serie.ticker, serie.quantidade, serie.custo) == ("XXXXA17", -1000, D(-498))
    ap = livro.apurar(date(2025, 1, 31))
    assert categoria(ap, 2025, 1, Categoria.COMUM).imposto == D("74.70")


def test_c3b_serie_ajustada_por_desdobramento_e_exercida() -> None:
    livro = Livro()
    venc = date(2025, 9, 19)
    livro.nota(date(2025, 8, 1), livro.op("C", "YYYYI405", 300, "1.20", mercado="CALL",
                                          vencimento=venc), custos="0.90")  # fmt: skip
    livro.evento(date(2025, 8, 12), "YYYYI405", INC, quantidade="300", para="YYYYI135",
                 quantidade_destino="900", natureza=SERIE, tipo_ativo=OPCAO)  # fmt: skip
    [serie] = livro.carteira(date(2025, 8, 12))
    assert (serie.quantidade, serie.custo, serie.custo / serie.quantidade) == (
        900,
        D("360.90"),
        D("0.401"),
    )
    livro.nota(date(2025, 9, 19), livro.op("C", "YYYY3", 900, "13.50", mercado="EXERC CALL",
                                           opcao="YYYYI135"), custos="3.60")  # fmt: skip
    [acao] = livro.carteira(date(2025, 9, 19))
    assert (acao.ticker, acao.quantidade, acao.custo) == ("YYYY3", 900, D("12514.50"))
    assert all(m.resultado == 0 for m in livro.movimentos(date(2025, 9, 19)))


def test_c3c_quantidade_truncada_pela_b3_fica_com_o_custo_inteiro() -> None:
    livro = Livro()
    livro.nota(
        date(2024, 12, 2),
        livro.op("C", "XXXXA180", 1000, "0.50", mercado="CALL", vencimento=date(2025, 1, 17)),
        custos="2.00",
    )
    livro.evento(date(2024, 12, 16), "XXXXA180", INC, quantidade="1000", para="XXXXA17",
                 quantidade_destino="1030", natureza=SERIE, tipo_ativo=OPCAO)  # fmt: skip
    [serie] = livro.carteira(date(2024, 12, 16))
    assert (serie.quantidade, serie.custo) == (1030, D(502))
    assert (serie.custo / serie.quantidade).quantize(D("0.00000001")) == D("0.48737864")
