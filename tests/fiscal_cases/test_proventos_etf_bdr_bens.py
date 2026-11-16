"""Proventos que ficavam fora do cálculo (item B10 do parecer do especialista de 16/11/2026).

- Rendimento distribuído por ETF de ações desde 01/01/2024: IRRF de 15% na data da distribuição
  (Lei 14.754/2023 arts. 18 II e 24 caput), definitivo para a pessoa física (art. 32 I): ficha de
  tributação exclusiva, código 06, pelo líquido. Antes de 2024 a IN RFB 1.585/2015 (arts. 24 a
  27) não trata da distribuição: só aviso.
- Rendimento distribuído por ETF de renda fixa: exclusiva na fonte na distribuição (Lei
  13.043/2014 art. 2º §4º; IN RFB 1.585/2015 arts. 29 III e 30 II): exclusiva 06, pelo líquido.
- Dividendo de BDR: não é isento (a Lei 9.249/1995 art. 10 só alcança PJ tributada no País);
  rendimento do exterior, carnê-leão pela leitura literal (Lei 7.713/1988 art. 8º; a Lei
  14.754/2023 art. 3º §1º I só alcança operações fora do País). Só aviso, sem cálculo.
- Dividendo pago em bens (ações ou cotas de outra empresa): rendimento isento pelo valor
  atribuído no aviso (Lei 9.249/1995 art. 10; linha 09). Custo do bem: o valor atribuído ou,
  quando menor, o valor corrente na entrega (IN RFB 1.585/2015 art. 58 §3º IV; Lei 7.713/1988
  art. 16 V); critério do COIR, conservador (parâmetro custo_dividendo_em_bens).

Nenhum deles entra na apuração mensal. Números conferidos à mão."""

from datetime import date

import pytest

from coir.domain.fiscal import Categoria, TipoAtivo, TipoEvento
from coir.domain.parametros import TiposAtivo
from coir.domain.proventos import LinhaB3, Provento, TipoProvento
from coir.engine.proventos import proventos
from coir.erros import MotorError
from coir.referencias import carregar_irpf
from tests.fiscal_cases.apoio import REFS, D, Livro, categoria, parametros
from tests.fiscal_cases.test_proventos import mov

# ETF de ações, ETF de renda fixa e BDR de exemplo (códigos fictícios)
TIPOS = TiposAtivo(tabela={**REFS.tipos.tabela, "XETF11": TipoAtivo.ETF,
                           "XERF11": TipoAtivo.ETF_RF, "XBDR34": TipoAtivo.BDR})  # fmt: skip
BENS = TipoEvento.DIVIDENDO_EM_BENS


def _proventos(*linhas: LinhaB3) -> tuple[list[Provento], list[str]]:
    avisos: list[str] = []
    return proventos(linhas, TIPOS, REFS.parametros, avisos), avisos


@pytest.mark.parametrize("movimentacao", ["Rendimento", "Dividendo"])
def test_rendimento_de_etf_de_acoes_desde_2024_e_exclusivo_pelo_liquido(movimentacao: str) -> None:
    """Caso B10-b: 1.000 cotas a 0,10 = bruto 100,00; IRRF de 15% = 15,00; líquido 85,00
    creditado pela B3. Exclusiva 06: 85,00 (Lei 14.754/2023 arts. 18 II, 24 e 32 I)."""
    [p], avisos = _proventos(mov(date(2025, 6, 10), movimentacao, "XETF11 - ETF X", "1000",
                                 "0.10", "85"))  # fmt: skip
    assert p.tipo is TipoProvento.RENDIMENTO_ETF
    assert (p.liquido, p.bruto, p.irrf, p.aliquota) == (D(85), D("100.00"), D("15.00"), D("0.150"))
    assert p.ano_declaracao == 2025 and not avisos


def test_rendimento_de_etf_de_acoes_antes_de_2024_fica_so_no_aviso() -> None:
    """Até 31/12/2023 a IN RFB 1.585/2015 arts. 24 a 27 não trata da distribuição: o COIR não
    classifica (parâmetro rendimento_etf_exclusivo falso até 2023)."""
    [p], avisos = _proventos(mov(date(2023, 12, 28), "Rendimento", "XETF11 - ETF X", "1000",
                                 "0.10", "85"))  # fmt: skip
    assert p.tipo is TipoProvento.FORA_DO_ESCOPO
    assert any("ETF de ações antes de 2024" in a for a in avisos)


def test_rendimento_de_etf_de_renda_fixa_e_exclusivo_em_qualquer_ano() -> None:
    """Lei 13.043/2014 art. 2º §4º; IN RFB 1.585/2015 arts. 29 III e 30 II. 500 cotas a 0,20 =
    bruto 100,00; líquido 85,00. Sem o preço na linha, o bruto fica igual ao líquido."""
    [a, b], _ = _proventos(
        mov(date(2022, 5, 5), "Rendimento", "XERF11 - ETF RF", "500", "0.20", "85"),
        mov(date(2025, 5, 5), "Rendimento", "XERF11 - ETF RF", "500", None, "85", i=2),
    )
    assert a.tipo is b.tipo is TipoProvento.RENDIMENTO_ETF_RF
    assert (a.liquido, a.bruto, a.irrf) == (D(85), D("100.00"), D("15.00"))
    assert (b.liquido, b.bruto, b.irrf) == (D(85), D(85), D(0))


def test_dividendo_de_bdr_so_avisa_com_a_leitura_do_carne_leao() -> None:
    """Lei 9.249/1995 art. 10 (só PJ tributada no País); Lei 7.713/1988 art. 8º. O COIR não
    calcula: faltam o imposto retido no exterior e a tabela progressiva."""
    [p], avisos = _proventos(mov(date(2025, 3, 3), "Dividendo", "XBDR34 - EMPRESA X", "10", "1",
                                 "10"))  # fmt: skip
    assert p.tipo is TipoProvento.FORA_DO_ESCOPO
    [aviso] = avisos
    assert "não é isento" in aviso and "carnê-leão" in aviso and "depositário" in aviso


@pytest.mark.parametrize("ano_calendario", [2024, 2025, 2026])
def test_linhas_da_exclusiva_06_para_os_rendimentos_de_etf(ano_calendario: int) -> None:
    """Programa IRPF, ficha Rendimentos sujeitos à tributação exclusiva/definitiva, código 06
    (rendimentos de aplicações financeiras), desde o exercício 2025 (ano-calendário 2024)."""
    mapa = carregar_irpf(ano_calendario)
    for chave in ("rendimento_etf", "rendimento_etf_rf"):
        assert mapa.exclusivos.linha(chave).codigo == "06"
        assert not mapa.isentos.tem(chave)


# ------------------------------------------------------------------ dividendo pago em bens
def _dividendo_em_bens(corrente: str) -> Livro:
    """Caso B10-d: a companhia ABCD3 paga parte do dividendo com 20 cotas do fundo XFII11, com
    valor atribuído de 100,00 cada (2.000,00), entregues em 10/06/2019."""
    livro = Livro()
    livro.nota(date(2019, 1, 7), livro.op("C", "ABCD3", 1000, "10"))
    livro.evento(date(2019, 6, 10), "ABCD3", BENS, valor="2000", para="XFII11",
                 quantidade_destino="20", valor_corrente=corrente)  # fmt: skip
    return livro


def _vender_as_cotas(livro: Livro) -> None:
    livro.nota(date(2019, 9, 2), livro.op("V", "XFII11", 20, "110", tipo=TipoAtivo.FII))


def test_dividendo_em_bens_entra_pelo_valor_atribuido_quando_o_de_mercado_nao_e_menor() -> None:
    """Valor de mercado na entrega 2.100,00, acima do atribuído (2.000,00): custo 2.000,00. Venda
    das 20 cotas a 110,00 = 2.200,00: ganho de FII 200,00, imposto de 20% = 40,00 (Lei 8.668/1993
    art. 18 II). A posição na companhia que pagou não muda."""
    livro = _dividendo_em_bens("2100")
    [cotas] = [p for p in livro.carteira(date(2019, 6, 10)) if p.ticker == "XFII11"]
    [acoes] = [p for p in livro.carteira(date(2019, 6, 10)) if p.ticker == "ABCD3"]
    assert (cotas.quantidade, cotas.custo) == (D(20), D(2000))
    assert (acoes.quantidade, acoes.custo) == (D(1000), D(10000))
    _vender_as_cotas(livro)
    ap = livro.apurar(date(2019, 9, 30))
    fii = categoria(ap, 2019, 9, Categoria.FII)
    assert (fii.resultado, fii.imposto) == (D(200), D("40.00"))
    assert not [a for a in ap.avisos if "valor de mercado na entrega" in a]


def test_dividendo_em_bens_fica_pelo_valor_de_mercado_quando_menor_com_aviso() -> None:
    """Leitura conservadora (IN RFB 1.585/2015 art. 58 §3º IV; Lei 7.713/1988 art. 16 V): mercado
    na entrega 1.900,00, abaixo do atribuído: custo 1.900,00. Venda por 2.200,00: ganho de FII
    300,00, imposto 60,00. O rendimento isento continua 2.000,00 (sai do evento, não do custo)."""
    livro = _dividendo_em_bens("1900")
    _vender_as_cotas(livro)
    ap = livro.apurar(date(2019, 9, 30))
    fii = categoria(ap, 2019, 9, Categoria.FII)
    assert (fii.resultado, fii.imposto) == (D(300), D("60.00"))
    [aviso] = [a for a in ap.avisos if "pago em XFII11" in a]
    assert "R$ 1.900,00" in aviso and "R$ 2.000,00" in aviso and "rendimento isento" in aviso
    [entrada] = [m for m in ap.movimentos if m.evento is BENS]
    assert (entrada.ticker, entrada.cv, entrada.quantidade, entrada.valor, entrada.resultado) == (
        "XFII11", "C", D(20), D(1900), D(0))  # fmt: skip


def test_parametro_devolve_o_custo_pelo_valor_atribuido() -> None:
    """custo_dividendo_em_bens = valor_atribuido (leitura A do parecer): custo 2.000,00 mesmo com
    o mercado a 1.900,00; ganho 200,00, imposto 40,00."""
    livro = _dividendo_em_bens("1900")
    _vender_as_cotas(livro)
    ap = livro.apurar(date(2019, 9, 30), parametros(custo_dividendo_em_bens="valor_atribuido"))
    assert categoria(ap, 2019, 9, Categoria.FII).imposto == D("40.00")


def test_dividendo_em_bens_nao_entra_na_apuracao_do_mes_da_entrega() -> None:
    """O dividendo é rendimento isento (Lei 9.249/1995 art. 10): sem resultado, sem venda no
    limite de R$ 20 mil e sem imposto em junho."""
    ap = _dividendo_em_bens("1900").apurar(date(2019, 6, 30))
    assert not [m for m in ap.meses if (m.ano, m.mes) == (2019, 6) and m.imposto_devido]
    assert all(m.resultado == 0 and m.venda_acao_a_vista == 0 for m in ap.movimentos)


def test_dividendo_em_bens_sem_o_valor_de_mercado_nao_e_calculado() -> None:
    livro = Livro()
    livro.evento(date(2019, 6, 10), "ABCD3", BENS, valor="2000", para="XFII11",
                 quantidade_destino="20")  # fmt: skip
    with pytest.raises(MotorError, match="valor de mercado"):
        livro.movimentos(date(2019, 12, 31))
