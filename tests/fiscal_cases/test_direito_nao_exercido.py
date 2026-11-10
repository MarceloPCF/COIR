"""Direito, recibo ou bônus de subscrição comprado que vence sem exercício nem venda (evento
DIREITO_NAO_EXERCIDO): o custo pago, com corretagem e emolumentos, vira perda em operações comuns
na data do fim do prazo de exercício, fora do limite de R$ 20 mil (IN RFB 1.585/2015 art. 60 §3º,
por analogia com a opção que vira pó; arts. 56 §3º, 58 e 64; SC Cosit 166/2021; Lei 6.404 art.
171 §4º). Parecer do especialista de 09/11/2026, seguido pelo usuário."""

from datetime import date

from coir.domain.fiscal import Categoria, TipoAtivo, TipoEvento
from tests.fiscal_cases.apoio import D, Livro, categoria

DIR = TipoAtivo.DIREITO
VENCEU = TipoEvento.DIREITO_NAO_EXERCIDO


def test_caso_a_compra_que_vence_vira_prejuizo_e_compensa_no_mes_seguinte() -> None:
    """Compra de 1.000 a 0,40 com 2,10 de custos (custo 402,10); prazo até 27/03/2026. Março:
    -402,10, sem vendas no limite. Abril: vende 1.000 PETR4 por 25.000 (custo 23.597,90), ganho
    1.402,10; base 1.000,00 e IR 150,00."""
    livro = Livro()
    livro.nota(date(2026, 3, 10), livro.op("C", "XPTO1", 1000, "0.40", tipo=DIR), custos="2.10")
    livro.evento(date(2026, 3, 27), "XPTO1", VENCEU, quantidade="1000", tipo_ativo=DIR)
    livro.nota(date(2026, 3, 2), livro.op("C", "PETR4", 1000, "23.5979"))
    livro.nota(date(2026, 4, 15), livro.op("V", "PETR4", 1000, "25"))
    ap = livro.apurar(date(2026, 4, 30))
    marco = categoria(ap, 2026, 3, Categoria.COMUM)
    assert (marco.resultado, marco.vendas_acoes, marco.imposto) == (D("-402.10"), 0, 0)
    abril = categoria(ap, 2026, 4, Categoria.COMUM)
    assert (abril.resultado, abril.imposto) == (D("1402.10"), D("150.00"))
    [baixa] = [m for m in ap.movimentos if m.evento is VENCEU]
    assert (baixa.quantidade, baixa.valor, baixa.custo_baixa) == (D(1000), 0, D("402.10"))
    assert "não exercido" in baixa.regra


def test_caso_b_direito_recebido_que_vence_nao_da_resultado() -> None:
    """Recebidos sem custo nunca entram no livro (só na venda): a baixa não acha posição."""
    livro = Livro()
    livro.nota(date(2026, 3, 2), livro.op("C", "PETR4", 100, "10"))
    livro.evento(date(2026, 3, 27), "XPTO1", VENCEU, quantidade="500", tipo_ativo=DIR)
    ap = livro.apurar(date(2026, 3, 31))
    assert not [m for m in ap.movimentos if m.evento is VENCEU]
    assert categoria(ap, 2026, 3, Categoria.COMUM).resultado == 0


def test_baixa_maior_que_a_posicao_so_baixa_o_comprado() -> None:
    """Comprou 50 por 50,00 e a B3 baixou 150 (100 recebidos sem custo): perda de 50,00."""
    livro = Livro()
    livro.nota(date(2026, 3, 10), livro.op("C", "XPTO1", 50, "1", tipo=DIR))
    livro.evento(date(2026, 3, 27), "XPTO1", VENCEU, quantidade="150", tipo_ativo=DIR)
    ap = livro.apurar(date(2026, 3, 31))
    assert categoria(ap, 2026, 3, Categoria.COMUM).resultado == D(-50)
    assert not [a for _, a in ap.avisos_titular if "continuam na posição" in a]


def test_baixa_menor_que_a_posicao_deixa_o_resto_com_aviso() -> None:
    livro = Livro()
    livro.nota(date(2026, 3, 10), livro.op("C", "XPTO1", 100, "1", tipo=DIR))
    livro.evento(date(2026, 3, 27), "XPTO1", VENCEU, quantidade="40", tipo_ativo=DIR)
    ap = livro.apurar(date(2026, 3, 31))
    assert categoria(ap, 2026, 3, Categoria.COMUM).resultado == D(-40)
    assert [a for _, a in ap.avisos_titular if "continuam na posição" in a]


def test_sem_quantidade_baixa_todo_o_saldo() -> None:
    """A linha "Direitos de Subscrição - Não Exercido" da B3 vem sem quantidade: o evento sem ela
    baixa todo o saldo comprado (80 por 0,50 = 40,00), sem aviso."""
    livro = Livro()
    livro.nota(date(2026, 3, 10), livro.op("C", "XPTO1", 80, "0.5", tipo=DIR))
    livro.evento(date(2026, 3, 27), "XPTO1", VENCEU, tipo_ativo=DIR)
    ap = livro.apurar(date(2026, 3, 31))
    assert categoria(ap, 2026, 3, Categoria.COMUM).resultado == D(-40)
    assert not [a for _, a in ap.avisos_titular if "continuam na posição" in a]
