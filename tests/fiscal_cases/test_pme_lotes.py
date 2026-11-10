"""Lotes isento e tributado de ações de PME (parecer do especialista de 08/11/2026).

Norma lida:
- Lei 13.043/2014 art. 16 caput (isenção até 31/12/2023, mercado à vista de bolsa) e art. 17:
  caput (aquisição a partir de 10/07/2014, incisos I a IV), §2º (perdas não compensam), §3º (fora
  do limite de R$ 20 mil), §5º (custo de quem comprou até 10/07/2014) e §6º ("As ações adquiridas
  e não alienadas até 31 de dezembro de 2023 terão seus custos de aquisição ajustados [...] ao
  maior valor entre o custo de aquisição efetivamente pago e a média do preço de fechamento,
  ponderada pelo volume negociado nos últimos 30 (trinta) pregões anteriores a 31 de dezembro de
  2023").
- IN RFB 1.585/2015 art. 58 caput (média ponderada), arts. 66 a 68 (repetem a lei) e art. 69-A
  (incluído pela IN RFB 1.916/2019): caput (sem isenção para as ações adquiridas a partir da data
  da oferta subsequente), §1º (inclusive preferência e bonificação), §2º (as anteriores permanecem
  isentas), §3º (IRRF do art. 63 na venda isenta é compensável) e §4º ("O investidor deverá
  controlar, de forma separada, as ações adquiridas antes da oferta pública subsequente, que ainda
  gozem da isenção, daquelas adquiridas a partir da oferta pública, que serão tributadas").
- SC Cosit 63/2020: o lote anterior continua isento mesmo com a companhia desenquadrada.

Sem norma, escolhas do parecer, confirmadas pelo usuário em 08/11/2026: lote isento sai primeiro
(parâmetro `pme_lote_ordem_venda`, valores `isento_primeiro` e `tributado_primeiro`); custódia
inicial com data na janela é isenta, fora da janela é tributada, em ambos os casos com aviso; o
ajuste do §6º compara o PM do lote isento, vale só para ele e entra em 01/01/2024, de modo que a
posição de 31/12/2023 fica pelo custo pago.

Contrato do motor:
- `Evento.pme_ate`: a mesma marca da compra (`OperacaoMotor.pme_ate`). O serviço a preenche pela
  data do evento (`IsencaoPME.compra_elegivel`). Sem marca, o evento que acrescenta ações
  (bonificação, subscrição, custódia) vai para o lote tributado.
- `OperacaoMotor.pme_media_fim` e `Evento.pme_media_fim`: média dos 30 pregões da coluna media_fim
  de data/isencao_pme.csv. None: sem ajuste.
- Movimento com `origem` "pme:fim" em 01/01/2024: valor = aumento do custo, sem resultado.
"""

from datetime import date
from decimal import Decimal

import pytest

from coir.domain.fiscal import Categoria, NaturezaEvento, OperacaoMotor, TipoEvento
from coir.erros import MotorError
from tests.fiscal_cases.apoio import REFS, D, Livro, categoria, mes, parametros

PME_ATE = date(2023, 12, 31)
MEDIA = D(44)  # média fictícia dos 30 pregões: a real vem de data/isencao_pme.csv


def _com_media(op: OperacaoMotor, media: Decimal = MEDIA) -> OperacaoMotor:
    return op.model_copy(update={"pme_media_fim": media})


def _marcar_evento(livro: Livro, media: Decimal | None = None) -> None:
    """Põe no último evento a marca que o serviço dará pela data (dentro da janela de aquisição)."""
    livro.eventos[-1] = livro.eventos[-1].model_copy(
        update={"pme_ate": PME_ATE, "pme_media_fim": media}
    )


def _dois_lotes(ticker: str = "PRIO3", qi: int = 100, pi: str = "20", qt: int = 100,
                pt: str = "30") -> Livro:  # fmt: skip
    """Compra isenta em 02/03/2020 e tributada em 01/04/2021 (depois da oferta de jan/2021)."""
    livro = Livro()
    livro.nota(date(2020, 3, 2), livro.op("C", ticker, qi, pi, pme_ate=PME_ATE))
    livro.nota(date(2021, 4, 1), livro.op("C", ticker, qt, pt))
    return livro


# ------------------------------------------------------------------- custódia inicial e a janela
@pytest.mark.parametrize(
    ("ticker", "dia", "isento"),
    [
        ("PRIO3", date(2013, 12, 31), False),  # antes de 10/07/2014 (art. 17 caput)
        ("PRIO3", date(2019, 12, 31), True),
        ("PRIO3", date(2020, 12, 31), True),
        ("PRIO3", date(2021, 12, 31), False),  # depois da oferta: não presume isenção
        ("AGRO3", date(2020, 12, 31), True),
        ("AGRO3", date(2021, 12, 31), False),
    ],
)
def test_data_da_custodia_decide_a_marca(ticker: str, dia: date, isento: bool) -> None:
    """A custódia inicial segue a janela do CSV pela data do evento. Dentro dela o COIR presume
    aquisição a partir de 10/07/2014 (com aviso); fora, não presume isenção (CTN art. 111 II)."""
    assert REFS.pme[ticker].compra_elegivel(dia) is isento


def test_cenario_a_custodia_e_compra_na_janela_formam_um_so_lote_isento() -> None:
    """Cenário A. Os dois são isentos: não há dois lotes, e o PM é a média do art. 58.
    Custódia 31/12/2019: 1.000 por 20.000. Compra 10/02/2020: 500 a 32 = 16.000.
    Lote: 1.500 por 36.000, PM 24,00.
    Venda 15/06/2020: 600 a 40 = 24.000; custo 600 x 24 = 14.400; ganho 9.600, isento, fora do
    limite (art. 67 §3º) e sem IRRF (companhia ainda enquadrada, art. 67 §6º).
    Sobra 900 por 21.600, PM 24,00."""
    livro = Livro()
    livro.evento(date(2019, 12, 31), "PRIO3", TipoEvento.CUSTODIA_INICIAL, "C", "1000", "20000")
    _marcar_evento(livro)
    livro.nota(date(2020, 2, 10), livro.op("C", "PRIO3", 500, "32", pme_ate=PME_ATE))
    livro.nota(date(2020, 6, 15), livro.op("V", "PRIO3", 600, "40"))
    ap = livro.apurar(date(2020, 6, 30))
    assert mes(ap, 2020, 6).isento_pme == D(9600)
    comum = categoria(ap, 2020, 6, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.resultado, comum.imposto) == (0, 0, 0)
    assert mes(ap, 2020, 6).darf == 0
    [p] = livro.carteira(date(2020, 6, 30))
    assert (p.quantidade, p.custo, p.pm) == (D(900), D(21600), D(24))


def test_custodia_depois_da_janela_fica_tributada() -> None:
    """Custódia 31/12/2021 sem marca: 1.000 por 40.000, PM 40.
    Venda 10/03/2022: 500 a 50 = 25.000 (acima de R$ 20 mil); custo 20.000; ganho 5.000.
    Imposto 15% = 750; IRRF 1,25 (0,005%); DARF 748,75, vence em 29/04/2022 (sexta)."""
    livro = Livro()
    livro.evento(date(2021, 12, 31), "PRIO3", TipoEvento.CUSTODIA_INICIAL, "C", "1000", "40000")
    livro.nota(date(2022, 3, 10), livro.op("V", "PRIO3", 500, "50"), irrf="1.25")
    ap = livro.apurar(date(2022, 3, 31))
    assert categoria(ap, 2022, 3, Categoria.COMUM).imposto == D(750)
    m = mes(ap, 2022, 3)
    assert (m.isento_pme, m.irrf_deduzido, m.darf) == (0, D("1.25"), D("748.75"))
    assert m.vencimento == date(2022, 4, 29)


def test_lote_anterior_a_2014_e_compra_na_janela() -> None:
    """Art. 17 caput: comprada antes de 10/07/2014 não é isenta; o custo informado na custódia já
    é o do §5º (o COIR não tem a média de 2014). Compra de 2015 é isenta.
    Tributado: 100 por 500 (PM 5). Isento: 200 a 8 = 1.600 (PM 8).
    Venda 15/03/2016: 250 a 12 = 3.000. Isento primeiro:
    200 isentas: 2.400 - 1.600 = 800, isento de PME.
    50 tributadas: 600 - 250 = 350; vendas no limite 600, logo isento comum (art. 59 I).
    Sobra 50 por 250."""
    livro = Livro()
    livro.evento(date(2013, 12, 31), "PRIO3", TipoEvento.CUSTODIA_INICIAL, "C", "100", "500")
    livro.nota(date(2015, 2, 3), livro.op("C", "PRIO3", 200, "8", pme_ate=PME_ATE))
    livro.nota(date(2016, 3, 15), livro.op("V", "PRIO3", 250, "12"))
    ap = livro.apurar(date(2016, 3, 31))
    assert mes(ap, 2016, 3).isento_pme == D(800)
    comum = categoria(ap, 2016, 3, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.isento, comum.resultado) == (D(600), D(350), 0)
    [p] = livro.carteira(date(2016, 3, 31))
    assert (p.quantidade, p.custo) == (D(50), D(250))


def test_parametro_tributado_primeiro_inverte_a_ordem() -> None:
    """Mesmas entradas do caso anterior com `pme_lote_ordem_venda=tributado_primeiro`:
    100 tributadas: 1.200 - 500 = 700; vendas no limite 1.200, isento comum.
    150 isentas: 1.800 - 150 x 8 = 600, isento de PME."""
    livro = Livro()
    livro.evento(date(2013, 12, 31), "PRIO3", TipoEvento.CUSTODIA_INICIAL, "C", "100", "500")
    livro.nota(date(2015, 2, 3), livro.op("C", "PRIO3", 200, "8", pme_ate=PME_ATE))
    livro.nota(date(2016, 3, 15), livro.op("V", "PRIO3", 250, "12"))
    ap = livro.apurar(date(2016, 3, 31), parametros(pme_lote_ordem_venda="tributado_primeiro"))
    assert mes(ap, 2016, 3).isento_pme == D(600)
    comum = categoria(ap, 2016, 3, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.isento) == (D(1200), D(700))


# -------------------------------------------------------------------- dois lotes na mesma posição
def test_cenario_b_vendas_parciais_com_pm_separado_e_irrf() -> None:
    """Cenário B (art. 69-A §§ 2º a 4º). Isento: 1.000 a 20 (PM 20). Tributado: 2.000 a 30 (PM 30).
    jun/2021: V 600 a 35 = 21.000, só isentas; custo 12.000; ganho 9.000 isento; vendas no limite
    0; a nota retém 1,05 mesmo assim (companhia desenquadrada; §3º: compensável). Sobra isento 400
    por 8.000, tributado 2.000 por 60.000, total 2.400 por 68.000.
    jul/2021: V 700 a 40 = 28.000. 400 isentas: 16.000 - 8.000 = 8.000 isento. 300 tributadas:
    12.000 - 9.000 = 3.000; vendas no limite 12.000, logo isento comum. IRRF 1,40; saldo 2,45.
    ago/2021: V 1.000 a 40 = 40.000, só tributadas; custo 30.000; ganho 10.000; imposto 1.500.
    IRRF 2,00; saldo 4,45, todo deduzido. DARF 1.495,55, vence em 30/09/2021. Sobra 700 por
    21.000, PM 30."""
    livro = _dois_lotes(qi=1000, qt=2000)
    livro.nota(date(2021, 6, 1), livro.op("V", "PRIO3", 600, "35"), irrf="1.05")
    livro.nota(date(2021, 7, 1), livro.op("V", "PRIO3", 700, "40"), irrf="1.40")
    livro.nota(date(2021, 8, 2), livro.op("V", "PRIO3", 1000, "40"), irrf="2.00")
    ap = livro.apurar(date(2021, 8, 31))

    jun = mes(ap, 2021, 6)
    assert (jun.isento_pme, jun.irrf_mes, jun.irrf_saldo) == (D(9000), D("1.05"), D("1.05"))
    assert categoria(ap, 2021, 6, Categoria.COMUM).vendas_acoes == 0
    [p] = livro.carteira(date(2021, 6, 30))
    assert (p.quantidade, p.custo) == (D(2400), D(68000))

    jul = mes(ap, 2021, 7)
    assert (jul.isento_pme, jul.imposto, jul.irrf_saldo) == (D(8000), 0, D("2.45"))
    comum = categoria(ap, 2021, 7, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.isento) == (D(12000), D(3000))

    ago = mes(ap, 2021, 8)
    assert (ago.isento_pme, ago.imposto, ago.irrf_deduzido) == (0, D(1500), D("4.45"))
    assert (ago.darf, ago.vencimento) == (D("1495.55"), date(2021, 9, 30))
    [p] = livro.carteira(date(2021, 8, 31))
    assert (p.quantidade, p.custo, p.pm) == (D(700), D(21000), D(30))


def test_cenario_b_agro3_usa_a_janela_propria() -> None:
    """AGRO3: isenta até 04/02/2021. Isento 100 a 20; tributado 100 a 30 (compra em 05/02/2021).
    V 150 a 40 = 6.000. 100 isentas: 4.000 - 2.000 = 2.000. 50 tributadas: 2.000 - 1.500 = 500,
    isento comum (vendas no limite 2.000)."""
    livro = Livro()
    livro.nota(date(2020, 3, 2), livro.op("C", "AGRO3", 100, "20", pme_ate=PME_ATE))
    livro.nota(date(2021, 2, 5), livro.op("C", "AGRO3", 100, "30"))
    livro.nota(date(2021, 6, 1), livro.op("V", "AGRO3", 150, "40"))
    ap = livro.apurar(date(2021, 6, 30))
    assert mes(ap, 2021, 6).isento_pme == D(2000)
    comum = categoria(ap, 2021, 6, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.isento) == (D(2000), D(500))


def test_venda_que_atravessa_os_lotes_reparte_os_custos() -> None:
    """Isento 100 a 8; tributado 100 a 8. V 200 a 10 = 2.000 com custos 20: líquido 1.980.
    Repartição pela quantidade: 990 de alienação em cada lote, ganho 190 em cada.
    Limite: só o bruto do lote tributado (1.000), logo isento comum 190."""
    livro = _dois_lotes(pi="8", pt="8")
    livro.nota(date(2021, 6, 1), livro.op("V", "PRIO3", 200, "10"), custos="20")
    ap = livro.apurar(date(2021, 6, 30))
    assert mes(ap, 2021, 6).isento_pme == D(190)
    comum = categoria(ap, 2021, 6, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.isento) == (D(1000), D(190))


def test_perda_do_lote_isento_nao_compensa_e_a_do_tributado_compensa() -> None:
    """Art. 17 §2º. Isento 100 a 50; tributado 100 a 50. V 150 a 40 = 6.000.
    100 isentas: 4.000 - 5.000 = -1.000: perda de PME, só para consulta.
    50 tributadas: 2.000 - 2.500 = -500: prejuízo comum a compensar."""
    livro = _dois_lotes(pi="50", pt="50")
    livro.nota(date(2021, 6, 1), livro.op("V", "PRIO3", 150, "40"))
    ap = livro.apurar(date(2021, 6, 30))
    m = mes(ap, 2021, 6)
    assert (m.isento_pme, m.perda_pme) == (0, D(-1000))
    comum = categoria(ap, 2021, 6, Categoria.COMUM)
    assert (comum.resultado, comum.prejuizo_a_compensar) == (D(-500), D(-500))


def test_limite_de_20_mil_soma_o_lote_tributado_e_nao_o_isento() -> None:
    """Art. 67 §3º: só o lote isento fica fora da soma. Isento 1.000 a 10; tributado 1.000 a 10;
    VALE3 100 a 50. No mesmo mês: V 1.500 PRIO3 a 20 e V 100 VALE3 a 120.
    1.000 isentas: 20.000 - 10.000 = 10.000 isento de PME.
    500 tributadas: 10.000 - 5.000 = 5.000. VALE3: 12.000 - 5.000 = 7.000.
    Vendas no limite: 10.000 + 12.000 = 22.000 > 20.000; tributável 12.000; imposto 1.800.
    IRRF 2,10 (0,005% de 42.000, retido também sobre o lote isento); DARF 1.797,90, vence em
    30/07/2021."""
    livro = _dois_lotes(qi=1000, pi="10", qt=1000, pt="10")
    livro.nota(date(2021, 4, 1), livro.op("C", "VALE3", 100, "50"))
    livro.nota(date(2021, 6, 1), livro.op("V", "PRIO3", 1500, "20"),
               livro.op("V", "VALE3", 100, "120"), irrf="2.10")  # fmt: skip
    ap = livro.apurar(date(2021, 6, 30))
    comum = categoria(ap, 2021, 6, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.isento, comum.resultado) == (D(22000), 0, D(12000))
    m = mes(ap, 2021, 6)
    assert (m.isento_pme, m.imposto, m.darf) == (D(10000), D(1800), D("1797.90"))
    assert m.vencimento == date(2021, 7, 30)


# ------------------------------------------------------------------------------ fim da isenção
def test_fim_da_isencao_ajusta_o_lote_isento_e_funde_os_lotes() -> None:
    """Art. 17 §6º e IN art. 68. Isento 1.000 a 20 = 20.000; tributado 500 a 50 = 25.000.
    31/12/2023: pelo custo pago, 1.500 por 45.000 (a declaração do ano usa o custo pago).
    01/01/2024: o isento vai ao maior entre 20.000 e 1.000 x 44 = 44.000 (aumento de 24.000); o
    tributado fica como está (art. 69-A). Fusão: 1.500 por 69.000, PM 46,00 (art. 58).
    V 01/03/2024: 1.000 a 50 = 50.000; custo 46.000; ganho 4.000; acima do limite; imposto 600.
    IRRF 2,50; DARF 597,50, vence em 30/04/2024. Sem o ajuste, o PM seria 30 e o imposto 3.000."""
    livro = Livro()
    livro.nota(date(2020, 3, 2), _com_media(livro.op("C", "PRIO3", 1000, "20", pme_ate=PME_ATE)))
    livro.nota(date(2021, 4, 1), livro.op("C", "PRIO3", 500, "50"))
    livro.nota(date(2024, 3, 1), livro.op("V", "PRIO3", 1000, "50"), irrf="2.50")

    [antes] = livro.carteira(date(2023, 12, 31))
    assert (antes.quantidade, antes.custo) == (D(1500), D(45000))
    [depois] = livro.carteira(date(2024, 1, 1))
    assert (depois.quantidade, depois.custo, depois.pm) == (D(1500), D(69000), D(46))

    [ajuste] = [m for m in livro.movimentos(date(2024, 1, 31)) if m.origem == "pme:fim"]
    assert (ajuste.data, ajuste.quantidade, ajuste.valor) == (date(2024, 1, 1), 0, D(24000))
    assert (ajuste.resultado, ajuste.quantidade_apos, ajuste.custo_apos) == (0, D(1500), D(69000))

    ap = livro.apurar(date(2024, 3, 31))
    comum = categoria(ap, 2024, 3, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.resultado) == (D(50000), D(4000))
    m = mes(ap, 2024, 3)
    assert (m.isento_pme, m.imposto, m.darf) == (0, D(600), D("597.50"))
    assert m.vencimento == date(2024, 4, 30)
    [p] = livro.carteira(date(2024, 3, 31))
    assert (p.quantidade, p.custo, p.pm) == (D(500), D(23000), D(46))


def test_fim_da_isencao_com_media_abaixo_do_custo_mantem_o_custo_pago() -> None:
    """Isento 1.000 a 50 = 50.000; média 44: maior valor é 50.000. Tributado 500 a 44 = 22.000.
    Fusão: 1.500 por 72.000, PM 48,00. V 01/03/2024: 1.000 a 50 = 50.000; custo 48.000; ganho
    2.000; imposto 300."""
    livro = Livro()
    livro.nota(date(2020, 3, 2), _com_media(livro.op("C", "PRIO3", 1000, "50", pme_ate=PME_ATE)))
    livro.nota(date(2021, 4, 1), livro.op("C", "PRIO3", 500, "44"))
    livro.nota(date(2024, 3, 1), livro.op("V", "PRIO3", 1000, "50"))
    ap = livro.apurar(date(2024, 3, 31))
    assert categoria(ap, 2024, 3, Categoria.COMUM).resultado == D(2000)
    assert mes(ap, 2024, 3).imposto == D(300)
    [p] = livro.carteira(date(2024, 3, 31))
    assert (p.quantidade, p.custo, p.pm) == (D(500), D(24000), D(48))


def test_venda_no_ultimo_pregao_e_ajuste_so_do_que_sobrou() -> None:
    """Isento 1.000 a 20. V 28/12/2023: 400 a 50 = 20.000; custo 8.000; ganho 12.000 isento (a
    venda ainda está no prazo). Sobra 600 por 12.000; ajuste: 600 x 44 = 26.400.
    V 01/03/2024: 600 a 50 = 30.000; ganho 3.600; imposto 540."""
    livro = Livro()
    livro.nota(date(2020, 3, 2), _com_media(livro.op("C", "PRIO3", 1000, "20", pme_ate=PME_ATE)))
    livro.nota(date(2023, 12, 28), livro.op("V", "PRIO3", 400, "50"))
    livro.nota(date(2024, 3, 1), livro.op("V", "PRIO3", 600, "50"))
    ap = livro.apurar(date(2024, 3, 31))
    assert mes(ap, 2023, 12).isento_pme == D(12000)
    assert categoria(ap, 2024, 3, Categoria.COMUM).resultado == D(3600)
    assert mes(ap, 2024, 3).imposto == D(540)


def test_lote_so_tributado_nao_recebe_o_ajuste() -> None:
    """Comprou só depois da oferta: não há lote isento nem ajuste (leitura do parecer: o §6º é do
    regime do art. 17, do qual o art. 69-A tira essas ações). V 01/03/2024: 500 a 50 = 25.000;
    custo 500 x 30 = 15.000; ganho 10.000; imposto 1.500; vence em 30/04/2024."""
    livro = Livro()
    livro.nota(date(2021, 4, 1), livro.op("C", "PRIO3", 500, "30"))
    livro.nota(date(2024, 3, 1), livro.op("V", "PRIO3", 500, "50"))
    ap = livro.apurar(date(2024, 3, 31))
    assert categoria(ap, 2024, 3, Categoria.COMUM).resultado == D(10000)
    m = mes(ap, 2024, 3)
    assert (m.imposto, m.darf, m.vencimento) == (D(1500), D(1500), date(2024, 4, 30))


# ------------------------------------------------------------------------- eventos e outros casos
def test_bonificacao_na_janela_entra_no_lote_isento() -> None:
    """Art. 17 IV e IN art. 67 IV. Isento 1.000 a 20 = 20.000; bonificação em 15/06/2020 de 100
    com custo 500: 1.100 por 20.500. V 01/09/2020: 1.100 a 30 = 33.000; ganho 12.500 isento."""
    livro = Livro()
    livro.nota(date(2020, 3, 2), livro.op("C", "PRIO3", 1000, "20", pme_ate=PME_ATE))
    livro.evento(date(2020, 6, 15), "PRIO3", TipoEvento.BONIFICACAO, "C", "100", "500")
    _marcar_evento(livro)
    livro.nota(date(2020, 9, 1), livro.op("V", "PRIO3", 1100, "30"))
    ap = livro.apurar(date(2020, 9, 30))
    assert mes(ap, 2020, 9).isento_pme == D(12500)
    assert categoria(ap, 2020, 9, Categoria.COMUM).vendas_acoes == 0


def test_bonificacao_depois_da_oferta_entra_no_lote_tributado() -> None:
    """IN art. 69-A §1º, expresso, ainda que a ação de origem seja isenta. Isento 1.000 a 20;
    bonificação em 15/06/2021 de 100 com custo 300: tributado 100 por 300.
    V 01/09/2021: 1.100 a 30. 1.000 isentas: 30.000 - 20.000 = 10.000. 100 tributadas: 3.000 -
    300 = 2.700; vendas no limite 3.000, logo isento comum."""
    livro = Livro()
    livro.nota(date(2020, 3, 2), livro.op("C", "PRIO3", 1000, "20", pme_ate=PME_ATE))
    livro.evento(date(2021, 6, 15), "PRIO3", TipoEvento.BONIFICACAO, "C", "100", "300")
    livro.nota(date(2021, 9, 1), livro.op("V", "PRIO3", 1100, "30"))
    ap = livro.apurar(date(2021, 9, 30))
    assert mes(ap, 2021, 9).isento_pme == D(10000)
    comum = categoria(ap, 2021, 9, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.isento) == (D(3000), D(2700))


def test_subscricao_depois_da_oferta_entra_no_lote_tributado() -> None:
    """IN art. 69-A §1º. Isento 1.000 a 20; subscrição em 01/02/2021 de 200 por 5.000.
    V 01/09/2021: 1.200 a 30. 1.000 isentas: 10.000. 200 tributadas: 6.000 - 5.000 = 1.000;
    vendas no limite 6.000, logo isento comum."""
    livro = Livro()
    livro.nota(date(2020, 3, 2), livro.op("C", "PRIO3", 1000, "20", pme_ate=PME_ATE))
    livro.evento(date(2021, 2, 1), "PRIO3", TipoEvento.SUBSCRICAO, "C", "200", "5000")
    livro.nota(date(2021, 9, 1), livro.op("V", "PRIO3", 1200, "30"))
    ap = livro.apurar(date(2021, 9, 30))
    assert mes(ap, 2021, 9).isento_pme == D(10000)
    comum = categoria(ap, 2021, 9, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.isento) == (D(6000), D(1000))


def test_desdobramento_reparte_a_quantidade_e_nao_o_custo() -> None:
    """Não é aquisição (IN art. 58 §7º II): cada lote guarda o custo. Isento 100 por 2.000;
    tributado 100 por 3.000. Desdobramento 1:5 (+800, repartido pela quantidade): 500 por 2.000 e
    500 por 3.000. V 500 a 10 = 5.000, só isentas: 5.000 - 2.000 = 3.000. Sobra 500 por 3.000."""
    livro = _dois_lotes()
    livro.evento(date(2021, 6, 15), "PRIO3", TipoEvento.DESDOBRAMENTO, "C", "800")
    livro.nota(date(2021, 9, 1), livro.op("V", "PRIO3", 500, "10"))
    ap = livro.apurar(date(2021, 9, 30))
    assert mes(ap, 2021, 9).isento_pme == D(3000)
    [p] = livro.carteira(date(2021, 9, 30))
    assert (p.quantidade, p.custo) == (D(500), D(3000))


def test_grupamento_reparte_a_quantidade_e_nao_o_custo() -> None:
    """Isento 1.000 por 2.000; tributado 1.000 por 3.000. Grupamento 10:1 (-1.800): 100 por 2.000 e
    100 por 3.000. V 100 a 50 = 5.000, só isentas: 3.000 isento. Sobra 100 por 3.000."""
    livro = _dois_lotes(qi=1000, pi="2", qt=1000, pt="3")
    livro.evento(date(2021, 6, 15), "PRIO3", TipoEvento.GRUPAMENTO, "V", "1800")
    livro.nota(date(2021, 9, 1), livro.op("V", "PRIO3", 100, "50"))
    ap = livro.apurar(date(2021, 9, 30))
    assert mes(ap, 2021, 9).isento_pme == D(3000)
    [p] = livro.carteira(date(2021, 9, 30))
    assert (p.quantidade, p.custo) == (D(100), D(3000))


def test_transferencia_entre_contas_nao_quebra_o_lote() -> None:
    """Art. 17 §1º exige ficar no depositário central; mudar de conta não tira de lá. Compra de 100
    a 20 na conta 1, transferida à conta 2, vendida lá a 30: ganho 1.000 isento."""
    livro = Livro()
    livro.nota(date(2020, 3, 2), livro.op("C", "PRIO3", 100, "20", pme_ate=PME_ATE))
    livro.evento(date(2020, 6, 1), "PRIO3", TipoEvento.TRANSFERENCIA, destino=2)
    livro.nota(date(2020, 9, 1), livro.op("V", "PRIO3", 100, "30"), conta=2)
    assert mes(livro.apurar(date(2020, 9, 30)), 2020, 9).isento_pme == D(1000)


def test_incorporacao_com_lote_isento_no_prazo_falha_alto() -> None:
    """Sem norma para a isenção na ação recebida, e sem caso real: continua falhando alto."""
    livro = Livro()
    livro.nota(date(2020, 3, 2), livro.op("C", "PRIO3", 100, "20", pme_ate=PME_ATE))
    livro.evento(date(2021, 6, 15), "PRIO3", TipoEvento.INCORPORACAO, quantidade="100",
                 para="BBBB3", quantidade_destino="100",
                 natureza=NaturezaEvento.FUSAO)  # fmt: skip
    with pytest.raises(MotorError, match="lote isento de PME"):
        livro.movimentos(date(2021, 6, 30))


def test_incorporacao_depois_do_fim_da_isencao_segue_normal() -> None:
    """Hoje `_origem_do_evento` falha porque `pme_ate` só é limpo na compra seguinte. Depois de
    31/12/2023 não há lote isento: B entra com o custo de A (100 por 2.000; sem média informada,
    sem ajuste)."""
    livro = Livro()
    livro.nota(date(2020, 3, 2), livro.op("C", "PRIO3", 100, "20", pme_ate=PME_ATE))
    livro.evento(date(2024, 3, 1), "PRIO3", TipoEvento.INCORPORACAO, quantidade="100", para="BBBB3",
                 quantidade_destino="100", natureza=NaturezaEvento.FUSAO)  # fmt: skip
    [p] = livro.carteira(date(2024, 3, 31))
    assert (p.ticker, p.quantidade, p.custo) == ("BBBB3", D(100), D(2000))


def test_day_trade_nao_toca_os_lotes_e_segue_a_data_da_compra() -> None:
    """IN art. 65 e art. 69-A caput. Compra e venda no dia, depois da oferta: 50 a 40 e 50 a 42;
    resultado 100, tributado como day trade, imposto 20; IRRF 1; DARF 19, vence em 30/06/2021. Os
    lotes (200 por 5.000) não mudam."""
    livro = _dois_lotes()
    livro.nota(date(2021, 5, 3), livro.op("C", "PRIO3", 50, "40"),
               livro.op("V", "PRIO3", 50, "42"), irrf_dt="1")  # fmt: skip
    ap = livro.apurar(date(2021, 5, 31))
    assert categoria(ap, 2021, 5, Categoria.DAY_TRADE).resultado == D(100)
    m = mes(ap, 2021, 5)
    assert (m.isento_pme, m.imposto, m.darf) == (0, D(20), D(19))
    assert m.vencimento == date(2021, 6, 30)
    [p] = livro.carteira(date(2021, 5, 31))
    assert (p.quantidade, p.custo) == (D(200), D(5000))


def test_exercicio_de_call_depois_da_oferta_entra_no_lote_tributado() -> None:
    """Data da aquisição é a do exercício (art. 69-A caput); o prêmio soma ao custo (art. 58).
    Isento 100 a 20. Call 100 a 1 (prêmio 100), exercida a 30 em 15/03/2021: tributado 100 por
    3.100. V 01/06/2021: 200 a 40. 100 isentas: 4.000 - 2.000 = 2.000. 100 tributadas: 4.000 -
    3.100 = 900; vendas no limite 4.000, logo isento comum."""
    livro = Livro()
    livro.nota(date(2020, 3, 2), livro.op("C", "PRIO3", 100, "20", pme_ate=PME_ATE))
    livro.nota(date(2021, 3, 1), livro.op("C", "PRIOC30", 100, "1", mercado="CALL",
                                          vencimento=date(2021, 3, 15)))  # fmt: skip
    livro.nota(date(2021, 3, 15), livro.op("C", "PRIO3", 100, "30", mercado="EXERC CALL",
                                           opcao="PRIOC30"))  # fmt: skip
    livro.nota(date(2021, 6, 1), livro.op("V", "PRIO3", 200, "40"))
    ap = livro.apurar(date(2021, 6, 30))
    assert mes(ap, 2021, 6).isento_pme == D(2000)
    comum = categoria(ap, 2021, 6, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.isento) == (D(4000), D(900))


@pytest.mark.parametrize(
    ("ticker", "ultima_isenta", "primeira_tributada"),
    [
        ("PRIO3", date(2021, 1, 29), date(2021, 2, 1)),  # RI: "a partir de 01/02/2021" não
        ("AGRO3", date(2021, 2, 4), date(2021, 2, 5)),  # novas ações negociadas em 05/02/2021
        ("SQIA3", date(2019, 9, 5), date(2019, 9, 6)),  # RI: "entre 10/07/2014 e 05/09/2019"
    ],
)
def test_data_de_corte_das_ofertas_subsequentes(
    ticker: str, ultima_isenta: date, primeira_tributada: date
) -> None:
    """IN RFB 1.585/2015 art. 69-A: compras a partir da oferta subsequente são tributadas; §2º: as
    anteriores continuam isentas. A data é a que a companhia publicou, quando há."""
    isencao = REFS.pme[ticker]
    assert isencao.compra_elegivel(ultima_isenta)
    assert not isencao.compra_elegivel(primeira_tributada)
    assert isencao.vendas_ate == PME_ATE


@pytest.mark.parametrize("ticker", ["CRDE3", "FRTA3", "GSHP3", "NUTR3", "PRNR3"])
def test_sem_oferta_subsequente_a_compra_vale_ate_o_fim_do_prazo(ticker: str) -> None:
    """Lista da B3 de 2014 e fontes secundárias: sem data de corte conhecida."""
    isencao = REFS.pme[ticker]
    assert isencao.compra_elegivel(date(2014, 7, 10)) and isencao.compra_elegivel(PME_ATE)
    assert not isencao.compra_elegivel(date(2014, 7, 9))
