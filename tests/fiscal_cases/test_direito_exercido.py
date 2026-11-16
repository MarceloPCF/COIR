"""Direito de subscrição exercido, recibo de subscrição e direito recebido na data do crédito.

- Custo das ações subscritas = preço de subscrição pago + custo médio dos direitos usados (Lei
  7.713/1988 art. 16 caput e §2º; IN RFB 1.585/2015 art. 58 caput; art. 60 II a e §1º, por
  analogia com a opção de compra exercida: o prêmio soma ao preço de exercício, sem resultado).
- O exercício não é alienação (Lei 6.404/1976 art. 171): sem resultado, fora do limite de R$ 20
  mil e sem IRRF.
- O recibo é ativo próprio até virar ação; a conversão é troca, com o custo integral (IN 1.585
  art. 58 §§4º e 6º, por analogia; Lei 7.713 art. 3º §3º).
- O direito recebido entra com custo zero na data do crédito (IN 1.585 art. 58 §7º, por analogia;
  Lei 7.713 art. 16 §4º) e compõe o custo médio dos direitos desde então (art. 58 caput).

Casos do parecer do especialista de 13/11/2026, conferidos à mão."""

from datetime import date

import pytest

from coir.domain.fiscal import Categoria, NaturezaEvento, TipoAtivo, TipoEvento
from tests.fiscal_cases.apoio import D, Livro, categoria

DIR = TipoAtivo.DIREITO
RECEBIDO = TipoEvento.DIREITO_RECEBIDO
EXERCICIO = TipoEvento.EXERCICIO_DIREITO
VENCEU = TipoEvento.DIREITO_NAO_EXERCIDO
RECIBO = NaturezaEvento.RECIBO_SUBSCRICAO


def _posicao(livro: Livro, dia: date, ticker: str) -> tuple[D, D]:
    achadas = [p for p in livro.carteira(dia) if p.ticker == ticker]
    return (achadas[0].quantidade, achadas[0].custo) if achadas else (D(0), D(0))


def _base() -> Livro:
    """1.000 ABCD3 por 20.000,00; em 05/03/2026 a companhia credita 200 direitos ABCD1."""
    livro = Livro()
    livro.nota(date(2026, 1, 5), livro.op("C", "ABCD3", 1000, "20"))
    livro.evento(date(2026, 3, 5), "ABCD1", RECEBIDO, "C", quantidade="200", tipo_ativo=DIR)
    return livro


def _compra_e_venda_de_direitos(livro: Livro) -> None:
    """10/03: compra 100 ABCD1 a 1,50 com 0,30 de custos (150,30): 300 direitos por 150,30, custo
    médio 0,501. 12/03: vende 100 a 2,00 com 0,20 de custos: 199,80 - 50,10 = 149,70."""
    livro.nota(date(2026, 3, 10), livro.op("C", "ABCD1", 100, "1.50", tipo=DIR), custos="0.30")
    livro.nota(date(2026, 3, 12), livro.op("V", "ABCD1", 100, "2", tipo=DIR), custos="0.20")


def _exercicio(livro: Livro, quantidade: str = "200", valor: str = "2000") -> None:
    livro.evento(date(2026, 3, 20), "ABCD1", EXERCICIO, quantidade=quantidade, valor=valor,
                 para="ABCD9", quantidade_destino=quantidade, tipo_ativo=DIR)  # fmt: skip


def _conversao(livro: Livro, quantidade: str = "200") -> None:
    livro.evento(date(2026, 4, 15), "ABCD9", TipoEvento.INCORPORACAO, quantidade=quantidade,
                 para="ABCD3", quantidade_destino=quantidade, natureza=RECIBO,
                 tipo_ativo=DIR)  # fmt: skip


def test_exercicio_parcial_com_direitos_recebidos_e_comprados() -> None:
    """B4-1. Depois da venda sobram 200 direitos por 100,20. Exercício de 200 pagando 2.000,00:
    recibo ABCD9 com 200 por 2.100,20. Conversão em 15/04: ABCD3 com 1.200 por 22.100,20. Março:
    ganho de 149,70 no direito, fora do limite; imposto 22,455, truncado 22,45."""
    livro = _base()
    _compra_e_venda_de_direitos(livro)
    _exercicio(livro)
    _conversao(livro)
    assert _posicao(livro, date(2026, 3, 19), "ABCD1") == (D(200), D("100.20"))
    assert _posicao(livro, date(2026, 3, 31), "ABCD1") == (0, 0)
    assert _posicao(livro, date(2026, 3, 31), "ABCD9") == (D(200), D("2100.20"))
    assert _posicao(livro, date(2026, 3, 31), "ABCD3") == (D(1000), D(20000))  # ainda sem as novas
    assert _posicao(livro, date(2026, 4, 30), "ABCD9") == (0, 0)
    assert _posicao(livro, date(2026, 4, 30), "ABCD3") == (D(1200), D("22100.20"))
    ap = livro.apurar(date(2026, 4, 30))
    marco = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (marco.resultado, marco.vendas_limite, marco.isento) == (D("149.70"), 0, 0)
    assert marco.imposto == D("22.45")
    assert categoria(ap, 2026, 4, Categoria.COMUM).resultado == 0  # exercício e conversão: nada
    saida, entrada = (m for m in ap.movimentos if m.evento is EXERCICIO)
    assert (saida.cv, saida.quantidade, saida.valor) == ("V", D(200), D("100.20"))
    assert (entrada.cv, entrada.quantidade, entrada.valor) == ("C", D(200), D("2100.20"))
    assert "sem resultado" in saida.regra and "valor pago" in entrada.regra


def test_direito_so_comprado_leva_o_custo_para_a_acao() -> None:
    """B4-2. Compra de 500 ABCD1 a 0,80 com 0,40 de custos: 400,40. Exercício a 10,00 (5.000,00):
    recibo por 5.400,40, depois 500 ABCD3 por 5.400,40. Venda de 500 a 12,00 em 20/05, a única do
    mês (6.000,00): ganho isento de 599,60, e não de 1.000,00."""
    livro = Livro()
    livro.nota(date(2026, 3, 10), livro.op("C", "ABCD1", 500, "0.80", tipo=DIR), custos="0.40")
    _exercicio(livro, "500", "5000")
    _conversao(livro, "500")
    livro.nota(date(2026, 5, 20), livro.op("V", "ABCD3", 500, "12"))
    assert _posicao(livro, date(2026, 4, 30), "ABCD3") == (D(500), D("5400.40"))
    assert _posicao(livro, date(2026, 4, 30), "ABCD1") == (0, 0)  # o direito não fica na carteira
    maio = categoria(livro.apurar(date(2026, 5, 31)), 2026, 5, Categoria.COMUM)
    assert (maio.isento, maio.resultado, maio.imposto) == (D("599.60"), 0, 0)


def test_recibo_da_propria_subscricao_vendido_antes_da_conversao() -> None:
    """B4-3. Continua o B4-1 sem a conversão: em 25/03 vende os 200 ABCD9 a 11,00 com 0,40 de
    custos: 2.199,60 - 2.100,20 = 99,40. Março: 149,70 + 99,40 = 249,10; imposto 37,365, truncado
    37,36. Nenhuma ação nova entra em ABCD3."""
    livro = _base()
    _compra_e_venda_de_direitos(livro)
    _exercicio(livro)
    livro.nota(date(2026, 3, 25), livro.op("V", "ABCD9", 200, "11", tipo=DIR), custos="0.40")
    marco = categoria(livro.apurar(date(2026, 3, 31)), 2026, 3, Categoria.COMUM)
    assert (marco.resultado, marco.imposto, marco.vendas_limite) == (D("249.10"), D("37.36"), 0)
    assert _posicao(livro, date(2026, 3, 31), "ABCD9") == (0, 0)
    assert _posicao(livro, date(2026, 3, 31), "ABCD3") == (D(1000), D(20000))


def test_recibo_comprado_em_bolsa_vira_acao_com_o_custo() -> None:
    """B4-4. Compra de 300 ABCD9 a 11,00 com 0,90 de custos: 3.300,90. Na conversão, 300 ABCD3
    por 3.300,90, sem resultado."""
    livro = Livro()
    livro.nota(date(2026, 3, 23), livro.op("C", "ABCD9", 300, "11", tipo=DIR), custos="0.90")
    _conversao(livro, "300")
    assert _posicao(livro, date(2026, 4, 30), "ABCD3") == (D(300), D("3300.90"))
    assert categoria(livro.apurar(date(2026, 4, 30)), 2026, 4, Categoria.COMUM).resultado == 0


def test_venda_da_acao_antes_da_conversao_usa_o_custo_medio_sem_as_novas() -> None:
    """B4-5. Exercício de 200 direitos recebidos pagando 2.000,00: recibo com 200 por 2.000,00.
    Em 30/03 vende 500 ABCD3 a 19,00 (9.500,00): o custo é 500 x 20,00 = 10.000,00, prejuízo de
    500,00 (mês abaixo do limite; o prejuízo fica para compensar). Depois da conversão: 700 por
    12.000,00. Com as ações novas contadas antes da hora, o custo médio seria 18,33 e a venda
    daria ganho de 333,33."""
    livro = _base()
    _exercicio(livro)
    livro.nota(date(2026, 3, 30), livro.op("V", "ABCD3", 500, "19"))
    _conversao(livro)
    ap = livro.apurar(date(2026, 4, 30))
    marco = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (marco.vendas_limite, marco.resultado) == (D(9500), D(-500))
    assert marco.prejuizo_a_compensar == D(-500)
    assert _posicao(livro, date(2026, 4, 30), "ABCD3") == (D(700), D(12000))


def test_direitos_que_sobram_e_vencem_viram_perda_pelo_custo_medio() -> None:
    """B5-1. Como o B4-1, mas os 200 direitos restantes vencem em 06/04 sem exercício. Março:
    +149,70, imposto 22,45. Abril: perda de 200 x 0,501 = 100,20. Lançado só na venda, o direito
    recebido daria +49,50 em março e nada em abril."""
    livro = _base()
    _compra_e_venda_de_direitos(livro)
    livro.evento(date(2026, 4, 6), "ABCD1", VENCEU, tipo_ativo=DIR)
    ap = livro.apurar(date(2026, 4, 30))
    marco = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (marco.resultado, marco.imposto) == (D("149.70"), D("22.45"))
    abril = categoria(ap, 2026, 4, Categoria.COMUM)
    assert (abril.resultado, abril.prejuizo_a_compensar) == (D("-100.20"), D("-100.20"))


def test_so_recebidos_e_todos_vendidos_da_o_mesmo_com_ou_sem_o_credito_lancado() -> None:
    """B5-2. 100 recebidos em 05/03 e vendidos a 0,50 com 0,10 de custos: ganho 49,90, com o
    crédito lançado ou só na venda."""
    com = Livro()
    com.evento(date(2026, 3, 5), "ZZZZ12", RECEBIDO, "C", quantidade="100", tipo_ativo=DIR)
    sem = Livro()
    for livro in (com, sem):
        livro.nota(date(2026, 3, 10), livro.op("V", "ZZZZ12", 100, "0.50", tipo=DIR), custos="0.10")
        marco = categoria(livro.apurar(date(2026, 3, 31)), 2026, 3, Categoria.COMUM)
        assert (marco.resultado, marco.imposto) == (D("49.90"), D("7.48"))


def test_direito_vendido_sem_o_credito_e_com_comprados_avisa() -> None:
    """Sem o evento do crédito, os recebidos entram com custo zero só na venda. Havendo direitos
    comprados, o custo médio fica maior do que o devido: o COIR avisa para lançar o crédito."""
    livro = Livro()
    livro.nota(date(2026, 3, 10), livro.op("C", "ABCD1", 100, "1.50", tipo=DIR), custos="0.30")
    livro.nota(date(2026, 3, 12), livro.op("V", "ABCD1", 300, "2", tipo=DIR))
    ap = livro.apurar(date(2026, 3, 31))
    assert [a for _, a in ap.avisos_titular if "Direito recebido" in a and "ABCD1" in a]


def test_exercicio_de_direitos_fora_da_posicao_entra_com_custo_zero_e_avisa() -> None:
    """Exercício de 200 sem nenhum direito no livro (recebidos e não lançados): o recibo custa só
    os 2.000,00 pagos, com aviso para conferir a movimentação da B3."""
    livro = Livro()
    _exercicio(livro)
    assert _posicao(livro, date(2026, 3, 31), "ABCD9") == (D(200), D(2000))
    ap = livro.apurar(date(2026, 3, 31))
    assert [a for _, a in ap.avisos_titular if "recebidos sem custo" in a]


@pytest.mark.parametrize(
    ("campos", "motivo"),
    [
        ({"quantidade": None}, "direitos exercidos"),
        ({"quantidade_destino": None}, "quantidade subscrita"),
        ({"valor": "0"}, "valor pago"),
    ],
)
def test_exercicio_incompleto_isola_o_ativo(campos: dict[str, str | None], motivo: str) -> None:
    """Evento sem um dos dados não é calculado no palpite: o direito e o recibo saem da apuração,
    com o motivo, e o resto da carteira segue."""
    livro = _base()
    dados: dict[str, str | None] = {"quantidade": "200", "quantidade_destino": "200",
                                    "valor": "2000", **campos}  # fmt: skip
    livro.evento(date(2026, 3, 20), "ABCD1", EXERCICIO, quantidade=dados["quantidade"],
                 valor=dados["valor"] or "0", para="ABCD9",
                 quantidade_destino=dados["quantidade_destino"], tipo_ativo=DIR)  # fmt: skip
    isolados = livro.apurar(date(2026, 3, 31)).isolados
    assert {i.ticker for i in isolados} == {"ABCD1", "ABCD9"}
    assert all(motivo in i.motivo for i in isolados)


def test_direito_recebido_com_custo_isola_o_ativo() -> None:
    """O direito pago vem na nota de corretagem; o evento é só para o recebido, de custo zero."""
    livro = Livro()
    livro.evento(date(2026, 3, 5), "ABCD1", RECEBIDO, "C", quantidade="200", valor="10",
                 tipo_ativo=DIR)  # fmt: skip
    [isolado] = livro.apurar(date(2026, 3, 31)).isolados
    assert isolado.ticker == "ABCD1" and "custo zero" in isolado.motivo
