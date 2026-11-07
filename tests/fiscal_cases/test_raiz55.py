"""RAIZ4 -> RAIZ55 (parecer do especialista fiscal de 07/11/2026; números conferidos à mão).

Fatos: B3 OC 022/2026-VPE: RAIZ4 (BRRAIZACNPR6) ajustado automaticamente para RAIZ55 (BRRAIZA01PA3),
PN negociadas como PNA, fator de cotação e lote padrão de 100; COTAHIST: último negócio de RAIZ4 em
30/10/2026 e primeiro de RAIZ55 em 03/11/2026. Nada muda na quantidade nem nos direitos.

Regra: não há alienação (Lei 7.713/1988 art. 3º §3º), e o custo e a quantidade continuam os mesmos.
Na nota, a quantidade é a de ações e o preço é por 100 ações; o custo é o valor impresso (IN RFB
1.585/2015 art. 58 caput). A PNA continua ação para a isenção de R$ 20 mil (art. 59 caput).
"""

from datetime import date
from decimal import Decimal

from coir import referencias
from coir.domain.fiscal import Categoria, OperacaoMotor, TipoAtivo
from tests.fiscal_cases.apoio import REFS, Livro, categoria

D = Decimal


def _livro() -> Livro:
    """O mesmo ativo com os dois códigos, como faz a troca de data/trocas_ticker.csv."""
    livro = Livro()
    livro.ativos["RAIZ55"] = livro.ativo("RAIZ4")
    return livro


def _lote_100(livro: Livro, cv: str, quantidade: int, preco: str, impresso: str) -> OperacaoMotor:
    """Linha da nota: `impresso` é o valor da linha, quantidade x preço / 100 ao centavo."""
    op = livro.op(cv, "RAIZ55", quantidade, preco, tipo=TipoAtivo.ACAO)
    return op.model_copy(update={"total": D(impresso)})


def test_quantidade_nao_muda_e_o_custo_e_o_valor_impresso() -> None:
    """Nota do caso real: 10 x 24,54 / 100 = 2,454, impresso 2,45. O custo é 2,45, e não 245,40."""
    livro = _livro()
    livro.nota(date(2026, 11, 4), _lote_100(livro, "C", 10, "24.54", "2.45"))
    [posicao] = livro.carteira(date(2026, 11, 4))
    assert (posicao.quantidade, posicao.custo) == (D(10), D("2.45"))


def test_pm_atravessa_a_troca_e_a_venda_conta_na_isencao() -> None:
    """Antes da troca: 100 RAIZ4 a 0,30 = 30,00. Depois: 100 RAIZ55 a 24,00 por lote = 24,00.
    Posição: 200 ações, custo 54,00, PM 0,27.
    Venda de 100 RAIZ55 a 30,00 por lote = 30,00: 30,00 - 27,00 = 3,00, isento: o limite conta o
    valor impresso, e a PNA é ação. Fica 100 com custo 27,00."""
    livro = _livro()
    livro.nota(date(2026, 10, 15), livro.op("C", "RAIZ4", 100, "0.30"))
    livro.nota(date(2026, 11, 3), _lote_100(livro, "C", 100, "24.00", "24.00"))
    livro.nota(date(2026, 11, 16), _lote_100(livro, "V", 100, "30.00", "30.00"))
    comum = categoria(livro.apurar(date(2026, 11, 30)), 2026, 11, Categoria.COMUM)
    assert (comum.vendas_acoes, comum.isento, comum.imposto) == (D("30.00"), D("3.00"), 0)
    [posicao] = livro.carteira(date(2026, 11, 30))
    assert (posicao.quantidade, posicao.custo) == (D(100), D(27))


def test_cadastro_troca_e_tipo() -> None:
    refs = referencias.carregar()
    assert refs.ticker_na_data("RAIZ4", date(2026, 10, 30)) == "RAIZ55"
    assert REFS.tipos.tipo("RAIZ55") is TipoAtivo.ACAO
