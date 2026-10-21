"""Proventos a partir da movimentação da B3 (casos fictícios conferidos à mão, 19/10/2026).

Fontes: Lei 9.249/1995 arts. 9º e 10 (JCP exclusivo; dividendos isentos); LC 224/2025 art. 8º e
art. 14 III (JCP a 17,5% pago ou creditado desde 01/01/2026, na data do pagamento ou do crédito);
Lei 11.033/2004 art. 3º III (rendimentos de FII isentos); IN RFB 1.585/2015 art. 74 e Lei
15.265/2025 arts. 20 e 21 (reembolso de BTC); IN 1.585 art. 58 §8º (restituição de capital
reduz o custo); Lei 15.270/2025 (Lei 9.250 art. 6º-A: IRRF de 10% sobre o total dos dividendos da
mesma PJ no mês acima de R$ 50 mil). Decisões do usuário de 19/10/2026: JCP declarado no ano do
crédito; reembolso de dividendos de 2026 isento por equivalência; BDR fora da v3.0.
"""

from datetime import date
from decimal import Decimal

import pytest

from coir.domain.b3 import LinhaMovimentacao
from coir.domain.fiscal import TipoAtivo
from coir.domain.parametros import TiposAtivo
from coir.domain.proventos import LinhaB3, Provento, TipoProvento
from coir.engine.proventos import proventos
from coir.erros import MotorError
from tests.fiscal_cases.apoio import REFS

D = Decimal
XP = "XP INVESTIMENTOS CCTVM S/A"


def mov(dia: date, movimentacao: str, produto: str, qtd: str, preco: str | None, valor: str | None,
        entrada: bool = True, i: int = 1) -> LinhaB3:  # fmt: skip
    codigo = produto.split(" - ")[0]
    return LinhaB3(id=i, titular=1, linha=LinhaMovimentacao(
        data=dia, entrada=entrada, movimentacao=movimentacao, produto=produto, codigo=codigo,
        instituicao=XP, quantidade=D(qtd), preco=D(preco) if preco else None,
        valor=D(valor) if valor else None))  # fmt: skip


# um FII de exemplo (a tabela real só tem os ativos operados)
TIPOS = TiposAtivo(tabela={**REFS.tipos.tabela, "KNRI11": TipoAtivo.FII})


def _proventos(*linhas: LinhaB3) -> tuple[list[Provento], list[str]]:
    avisos: list[str] = []
    return proventos(linhas, TIPOS, REFS.parametros, avisos), avisos


def test_p1_jcp_de_2025_a_15_por_cento() -> None:
    """1.000 ações x R$ 0,50 bruto, pago em 19/12/2025: líquido 425,00; bruto 500,00 e IRRF 75,00
    (15%); declarado em 2025, código 10 pelo líquido."""
    [p], _ = _proventos(mov(date(2025, 12, 19), "Juros Sobre Capital Próprio", "ITSA4 - ITAUSA",
                            "1000", "0.5", "425"))  # fmt: skip
    assert (p.tipo, p.liquido, p.bruto, p.irrf, p.aliquota, p.ano_declaracao) == (
        TipoProvento.JCP,
        D("425"),
        D("500.00"),
        D("75.00"),
        D("0.15"),
        2025,
    )


def test_p2_jcp_pago_em_2026_creditado_em_2026_e_em_2025() -> None:
    """P2a: 2.000 x R$ 0,40 = 800,00 bruto, pago em 31/03/2026 com 17,5%: líquido 660,00; IRRF
    140,00; declarado em 2026. P2b: mesmo bruto creditado em 22/12/2025 e pago em 20/02/2026 com
    15% retidos: líquido 680,00 -> declarado em 2025 (ano do crédito), bruto 800,00, IRRF 120,00,
    e a receber em 31/12/2025 (Bens e Direitos 99.07). Pela data do pagamento daria 824,24."""
    [a, b], _ = _proventos(
        mov(date(2026, 3, 31), "Juros Sobre Capital Próprio", "ITSA4 - ITAUSA", "2000", "0.4",
            "660"),
        mov(date(2026, 2, 20), "Juros Sobre Capital Próprio", "BBAS3 - BANCO DO BRASIL", "2000",
            "0.4", "680", i=2),
    )  # fmt: skip
    assert (a.bruto, a.irrf, a.aliquota, a.ano_declaracao) == (D("800.00"), D("140.00"),
                                                                D("0.175"), 2026)  # fmt: skip
    assert (b.bruto, b.irrf, b.aliquota, b.ano_declaracao) == (D("800.00"), D("120.00"),
                                                                D("0.15"), 2025)  # fmt: skip
    assert b.a_receber_em_31_12 and not a.a_receber_em_31_12


def test_p1b_bruto_estimado_arredonda_os_centavos() -> None:
    """Líquido 10,00 a 15%: 11,7647... -> 11,76 (estimativa; vale o informe)."""
    [p], _ = _proventos(mov(date(2025, 5, 2), "Juros Sobre Capital Próprio", "ITSA4 - ITAUSA",
                            "10", None, "10"))  # fmt: skip
    assert (p.bruto, p.irrf) == (D("11.76"), D("1.76"))


def test_p3_lei_15270_dividendos_da_mesma_empresa_no_mes() -> None:
    """08/04/2026 ABCD4 30.000,00 e 22/04/2026 ABCD3 19.500,00 líquidos (55.000 bruto, 5.500
    retidos no segundo): soma líquida 49.500 > 45.000 = "possível retenção". Em 05/05 a soma do
    mês é outra; JCP da mesma empresa não soma."""
    linhas = (
        mov(date(2026, 4, 8), "Dividendo", "ITSA4 - ITAUSA", "10000", "3", "30000"),
        mov(date(2026, 4, 22), "Dividendo", "ITSA3 - ITAUSA", "5000", "5", "19500", i=2),
        mov(date(2026, 4, 25), "Juros Sobre Capital Próprio", "ITSA4 - ITAUSA", "1000", "20",
            "16500", i=3),
        mov(date(2026, 5, 5), "Dividendo", "ITSA4 - ITAUSA", "10000", "4", "40000", i=4),
    )  # fmt: skip
    lista, avisos = _proventos(*linhas)
    assert [p.tipo for p in lista].count(TipoProvento.DIVIDENDO) == 3
    [aviso] = [a for a in avisos if "Lei 15.270" in a]
    assert aviso.startswith("04/2026: dividendos de ITSA somam 49500") and "possível" in aviso
    _, avisos2 = _proventos(*linhas, mov(date(2026, 4, 28), "Dividendo", "ITSA4 - ITAUSA",
                                             "1000", "1", "1000", i=5))  # fmt: skip
    assert any("acima do limite" in a for a in avisos2)


def test_lei_15270_nao_vale_antes_de_2026() -> None:
    _, avisos = _proventos(mov(date(2025, 4, 8), "Dividendo", "ITSA4 - ITAUSA", "1", "1", "90000"))
    assert not [a for a in avisos if "Lei 15.270" in a]


def test_p4_reembolso_de_btc_isento_com_aviso_desde_2026() -> None:
    """Doador de 1.000 ações: reembolso de dividendo de 1.200,00. Em 2025 isento (IN 1.585 art.
    74); em 2026 isento por equivalência (Lei 15.265/2025 omissa), com aviso; o PM não muda."""
    [a, b], _ = _proventos(
        mov(date(2025, 6, 2), "Reembolso", "ITSA4 - ITAUSA", "1000", "1.2", "1200"),
        mov(date(2026, 6, 2), "Reembolso", "ITSA4 - ITAUSA", "1000", "1.2", "1200", i=2),
    )  # fmt: skip
    assert a.tipo is b.tipo is TipoProvento.REEMBOLSO_BTC
    assert "sem identificar" in a.observacao and "equivalência" in b.observacao
    assert a.reembolso_de == b.reembolso_de == ""  # sem provento do dia para comparar


def test_p5_restituicao_de_capital_nao_e_rendimento() -> None:
    [p], _ = _proventos(mov(date(2024, 12, 18), "Restituição de Capital", "SYNE3 - SYN", "1000",
                            "2", "2000"))  # fmt: skip
    assert p.tipo is TipoProvento.RESTITUICAO_CAPITAL and "reduz o custo" in p.observacao


def test_dividendo_de_bdr_e_rendimento_em_acao_ficam_fora_com_aviso() -> None:
    lista, avisos = _proventos(
        mov(date(2025, 3, 3), "Dividendo", "MMMC34 - 3M", "10", "1", "10"),
        mov(date(2025, 3, 4), "Rendimento", "BBAS3 - BANCO DO BRASIL", "10", "0.1", "1", i=2),
        mov(date(2025, 3, 5), "Rendimento", "KNRI11 - KINEA", "10", "1", "10", i=3),
    )  # fmt: skip
    assert [p.tipo for p in lista] == [TipoProvento.FORA_DO_ESCOPO, TipoProvento.FORA_DO_ESCOPO,
                                       TipoProvento.RENDIMENTO_FII]  # fmt: skip
    assert any("BDR" in a for a in avisos) and any("rendimento em ACAO" in a for a in avisos)


def test_transferidos_se_anulam_e_par_aberto_avisa() -> None:
    linhas = (
        mov(date(2025, 3, 3), "Dividendo - Transferido", "ITSA4 - ITAUSA", "10", "1", "10"),
        mov(date(2025, 3, 3), "Dividendo - Transferido", "ITSA4 - ITAUSA", "10", "1", "10",
            entrada=False, i=2),
        mov(date(2025, 3, 4), "Dividendo - Transferido", "ITSA4 - ITAUSA", "5", "1", "5", i=3),
    )  # fmt: skip
    lista, avisos = _proventos(*linhas)
    assert lista == [] and any("sem o par" in a for a in avisos)


def test_debito_de_provento_fora_de_transferencia_falha_alto() -> None:
    with pytest.raises(MotorError, match="a débito"):
        _proventos(mov(date(2025, 3, 3), "Dividendo", "ITSA4 - ITAUSA", "1", "1", "1",
                       entrada=False))  # fmt: skip


def test_atualizacao_do_provento_pela_selic_e_tributacao_exclusiva() -> None:
    """BBSE3 em 02/03/2026: dividendo de R$ 2,54996501627 por ação e "Rendimento" (atualização
    pela Selic do balanço ao pagamento) de R$ 0,05684162552 por ação, pago no mesmo dia (B3). Para
    100 ações: bruto 5,684 -> 5,68; retidos 22,5% como renda fixa (Lei 8.981/1995 art. 65 §4º c;
    Lei 11.033/2004 art. 1º) -> líquido 4,40, IRRF 1,28. Não é dividendo: fora da Lei 15.270."""
    lista, avisos = _proventos(
        mov(date(2026, 3, 2), "Dividendo", "BBSE3 - BB SEGURIDADE", "100", "2.54996501627",
            "255.00"),
        mov(date(2026, 3, 2), "Rendimento", "BBSE3 - BB SEGURIDADE", "100", "0.05684162552",
            "4.40", i=2),
    )  # fmt: skip
    dividendo, atualizacao = lista
    assert dividendo.tipo is TipoProvento.DIVIDENDO
    assert (atualizacao.tipo, atualizacao.bruto, atualizacao.irrf, atualizacao.liquido) == (
        TipoProvento.ATUALIZACAO,
        D("5.68"),
        D("1.28"),
        D("4.40"),
    )
    assert atualizacao.aliquota == D("0.225")
    assert not avisos


def _jcp(dia: date, preco: str, valor: str, credito: date | None = None) -> LinhaB3:
    item = mov(dia, "Juros Sobre Capital Próprio", "ITSA4 - ITAUSA", "1000", preco, valor)
    return item.model_copy(update={"credito": credito, "credito_origem": "B3" if credito else ""})


def test_jcp_creditado_no_ano_anterior_com_a_mesma_aliquota() -> None:
    """Caso C do especialista (F8, 20/10/2026): crédito em 16/12/2024, pago em 10/03/2025, 15% nos
    dois anos (1.000 x 0,50 bruto, líquido 425,00). Sem a data do crédito o motor o põe em 2025
    (não há mudança de alíquota para perceber); com ela (pacote da B3 ou informe), declarado em
    2024 e a receber em 31/12/2024 (P&R IRPF 2026 pergunta 320; Lei 9.249/1995 art. 9º §2º)."""
    [sem], _ = _proventos(_jcp(date(2025, 3, 10), "0.5", "425"))
    assert sem.ano_declaracao == 2025
    [p], _ = _proventos(_jcp(date(2025, 3, 10), "0.5", "425", date(2024, 12, 16)))
    assert (p.ano_declaracao, p.aliquota, p.bruto, p.credito, p.a_receber_em_31_12) == (
        2024, D("0.15"), D("500.00"), date(2024, 12, 16), True)  # fmt: skip
    assert "creditado em 16/12/2024 (B3)" in p.observacao


def test_jcp_creditado_em_2025_e_pago_em_2026_usa_a_aliquota_do_credito() -> None:
    """LC 224/2025 art. 8º: 17,5% para JCP pago ou creditado desde 01/01/2026, o que vier
    primeiro. Creditado em 20/12/2025 e pago em 15/04/2026: 15%, declarado em 2025."""
    [p], _ = _proventos(_jcp(date(2026, 4, 15), "0.5", "425", date(2025, 12, 20)))
    assert (p.ano_declaracao, p.aliquota) == (2025, D("0.15"))


def test_credito_depois_do_pagamento_falha_alto() -> None:
    with pytest.raises(MotorError, match="depois do pagamento"):
        _proventos(_jcp(date(2025, 3, 10), "0.5", "425", date(2025, 3, 11)))


def test_reembolso_de_jcp_identificado_pelo_valor_por_acao() -> None:
    """Parecer do especialista (21/10/2026): o reembolso é a quantidade emprestada vezes o valor
    por ação do provento do dia, menos o IR do JCP (Lei 13.043/2014 art. 7º §1º II). JCP de
    0,520530743 por ação pago em 14/03/2025 (pacote da B3), 1.000 ações emprestadas:
    1.000 x 0,520530743 x 0,85 = 442,45 (não é dividendo de 2,141847479: 206,57 ações). Isento,
    no ano do recebimento, sem 99.07; desde 2026, Lei 15.265/2025 art. 21 I."""
    publicos = {("VALE", date(2025, 3, 14)): [("JCP", D("0.520530743")),
                                               ("DIVIDENDO", D("2.141847479"))]}  # fmt: skip
    linha = mov(date(2025, 3, 14), "Reembolso", "VALE3 - VALE", "0", None, "442.45")
    avisos: list[str] = []
    [p] = proventos([linha], TIPOS, REFS.parametros, avisos, publicos)
    assert (p.tipo, p.reembolso_de, p.ano_declaracao, p.a_receber_em_31_12) == (
        TipoProvento.REEMBOLSO_BTC, "JCP", 2025, False)  # fmt: skip
    assert "reembolso de JCP de 1000 ações emprestadas" in p.observacao


def test_reembolso_de_dividendo_pela_movimentacao_e_lei_15270() -> None:
    """Reembolso de dividendo identificado pelo dividendo do mesmo dia na própria movimentação:
    dividendo de 50,00 por ação; reembolso de 25.000,00 = 500 ações emprestadas inteiras (como JCP
    a 15% seriam 588,24 ações: não). Com outro dividendo de 5.000,00 no mês, em 2026 o aviso da
    Lei 15.270 soma os três (55.000,00 no mês, acima de R$ 50 mil; decisão do usuário de
    21/10/2026)."""
    lista, avisos = _proventos(
        mov(date(2026, 4, 10), "Dividendo", "ITSA4 - ITAUSA", "500", "50", "25000"),
        mov(date(2026, 4, 10), "Reembolso", "ITSA4 - ITAUSA", "0", None, "25000", i=2),
        mov(date(2026, 4, 20), "Dividendo", "ITSA4 - ITAUSA", "100", "50", "5000", i=3),
    )  # fmt: skip
    [reembolso] = [p for p in lista if p.tipo is TipoProvento.REEMBOLSO_BTC]
    assert reembolso.reembolso_de == "DIVIDENDO"
    assert any("ITSA somam 55000" in a and "acima do limite" in a for a in avisos)
