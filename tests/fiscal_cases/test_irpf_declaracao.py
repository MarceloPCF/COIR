"""Declaração anual (F8): Bens e Direitos, rendimentos e fichas mês a mês que o COIR preenche.

Casos do especialista fiscal (parecer de 20/10/2026), conferidos à mão; custos zero.

Fontes: P&R IRPF 2026 perguntas 707, 720 (um item por ação, total das corretoras, custo médio) e
715; IN RFB 1.585/2015 arts. 58 §1º (bonificação: custo atribuído), 63 §8º III (saldo de IRRF de
dezembro na declaração) e 74; Lei 11.033/2004 art. 3º I (código 20); Lei 9.249/1995 arts. 9º e 10;
Manual do MIR (Bens e Direitos, rendimentos do capital); data/irpf/2026.yaml.
"""

from datetime import date, datetime

from coir import referencias
from coir.domain.fiscal import TipoAtivo, TipoEvento, Visao
from coir.domain.proventos import Provento, TipoProvento
from coir.domain.relatorio import DadosRelatorio, InfoAtivo
from coir.reports.irpf import NEGATIVO_ANTERIOR, PREJUIZO, declaracao, fichas
from tests.fiscal_cases.apoio import REFS, D, Livro

ITSA = InfoAtivo(emissor="ITSA", cnpj="61532644000115", razao_social="ITAUSA S.A.",
                 especificacao="PN N1", tipo_b3="COMPANHIA")  # fmt: skip
VALE = InfoAtivo(emissor="VALE", cnpj="33592510000154", razao_social="VALE S.A.",
                 especificacao="ON NM", tipo_b3="COMPANHIA")  # fmt: skip
ITUB = InfoAtivo(emissor="ITUB", cnpj="60872504000123", razao_social="ITAU UNIBANCO S.A.",
                 especificacao="PN N1", tipo_b3="COMPANHIA")  # fmt: skip
PETR = InfoAtivo(emissor="PETR", cnpj="33000167000101", razao_social="PETROBRAS",
                 especificacao="PN N2", tipo_b3="COMPANHIA")  # fmt: skip
XPML = InfoAtivo(emissor="XPML", cnpj="28757546000100", razao_social="XP MALLS FII",
                 especificacao="CI", tipo_b3="FII")  # fmt: skip
INFO = {"ITSA4": ITSA, "VALE3": VALE, "ITUB3": ITUB, "ITUB4": ITUB, "PETR4": PETR,
        "XPML11": XPML, "AAAA3": VALE, "CCCC3": VALE}  # fmt: skip


def _dados(livro: Livro, ano: int, proventos: tuple[Provento, ...] = ()) -> DadosRelatorio:
    fim, anterior = date(ano, 12, 31), date(ano - 1, 12, 31)
    fiscal = tuple(livro.carteira(fim))
    return DadosRelatorio(
        titular=1, nome="Contribuinte Ficticio", cpf="00000000000", ano=ano,
        gerado_em=datetime(2026, 10, 20), codigo_darf="6015", apuracao=livro.apurar(fim),
        ate=fim, data_carteira=fim, visao=Visao.TITULAR, carteira=fiscal,
        carteira_fiscal=fiscal, carteira_fiscal_anterior=tuple(livro.carteira(anterior)),
        contas={}, notas=(), proventos=proventos, irpf=referencias.carregar_irpf(ano),
        info_ativos=INFO, custodia=tuple(livro.carteira(fim, Visao.CORRETORA)),
        custodia_anterior=tuple(livro.carteira(anterior, Visao.CORRETORA)),
    )  # fmt: skip


def _provento(dia: date, codigo: str, tipo: TipoProvento, liquido: str,
              ano: int | None = None) -> Provento:  # fmt: skip
    return Provento(data=dia, titular=1, codigo=codigo, empresa=codigo, instituicao="XP",
                    tipo=tipo, quantidade=D("100"), liquido=D(liquido), bruto=D(liquido),
                    irrf=D("0"), aliquota=D("0"), ano_declaracao=ano or dia.year,
                    origem="mov:1")  # fmt: skip


def test_um_bem_por_acao_somando_corretoras_e_codigo_20() -> None:
    """Caso A (P&R 720): 01/03/2024 XP C 100 ITSA4 a 10,00; 03/06/2024 Clear C 100 a 12,00;
    10/02/2025 XP V 50 a 13,00; 01/08/2025 Clear C 50 a 14,00. Um item 03.01: 31/12/2024 =
    2.200,00 (200 a PM 11); 31/12/2025 = 2.350,00 (200 a PM 11,75), custódia XP 50 e Clear 150.
    Isentos 20 = 100,00 (50 x 2,00; vendas de fevereiro de 650)."""
    livro = Livro()
    livro.nota(date(2024, 3, 1), livro.op("C", "ITSA4", 100, "10"), corretora="XP", conta=1)
    livro.nota(date(2024, 6, 3), livro.op("C", "ITSA4", 100, "12"), corretora="CLEAR", conta=2)
    livro.nota(date(2025, 2, 10), livro.op("V", "ITSA4", 50, "13"), corretora="XP", conta=1)
    livro.nota(date(2025, 8, 1), livro.op("C", "ITSA4", 50, "14"), corretora="CLEAR", conta=2)
    f = fichas(_dados(livro, 2025))
    assert f is not None and len(f.bens) == 1
    b = f.bens[0]
    assert (b.grupo, b.codigo, b.ticker, b.cnpj, b.anterior, b.atual) == (
        "03", "01", "ITSA4", "61532644000115", D("2200"), D("2350"))  # fmt: skip
    assert b.discriminacao == (
        "200 ações PN de ITAUSA S.A., código ITSA4, custodiadas em CLEAR: "
        "150; XP: 50. Custo médio R$ 11,75 (IN RFB 1.585/2015 art. 58)."
    )
    [r] = [r for r in f.rendimentos if r.codigo == "20"]
    assert (r.ficha, r.cnpj, r.valor) == ("isentos", None, D("100.00"))


def test_bonificacao_codigo_18_e_custo_no_bem() -> None:
    """Caso I: 1.000 AAAA3 a 20,00; bonificação de 10% com custo atribuído de 5,00 por ação (data
    ex 02/05/2025): 03.01 em 31/12/2025 = 1.100 ações, 20.500,00; isentos 18 = 500,00 com o CNPJ
    da empresa (IN 1.585 art. 58 §1º)."""
    livro = Livro()
    livro.nota(date(2024, 5, 2), livro.op("C", "AAAA3", 1000, "20"))
    livro.evento(date(2025, 5, 2), "AAAA3", TipoEvento.BONIFICACAO, cv="C", quantidade="100",
                 valor="500")  # fmt: skip
    f = fichas(_dados(livro, 2025))
    assert f is not None
    assert (f.bens[0].anterior, f.bens[0].atual) == (D("20000"), D("20500"))
    [r] = [r for r in f.rendimentos if r.codigo == "18"]
    assert (r.cnpj, r.valor) == ("33592510000154", D("500"))


def test_proventos_por_cnpj_e_jcp_a_receber() -> None:
    """Casos J e C: dividendos PETR4 100 + 150 + 50 e ITUB3 20 + ITUB4 80 (mesmo CNPJ): código 09
    com PETR 300,00 e Itaú 100,00; XPML11 110,40 e reembolso de BTC de VALE3 35,00 no 99; JCP
    creditado em 2025 e pago em 10/03/2026 (líquido 425,00): código 10 = 425,00 em 2025 e 99.07
    com 0,00 em 31/12/2024 e 425,00 em 31/12/2025; em 2026 o 99.07 vai de 425,00 a 0,00."""
    dv, jcp = TipoProvento.DIVIDENDO, TipoProvento.JCP
    proventos = (
        _provento(date(2025, 3, 1), "PETR4", dv, "100"),
        _provento(date(2025, 6, 1), "PETR4", dv, "150"),
        _provento(date(2025, 9, 1), "PETR4", dv, "50"),
        _provento(date(2025, 4, 1), "ITUB3", dv, "20"),
        _provento(date(2025, 4, 1), "ITUB4", dv, "80"),
        _provento(date(2025, 5, 15), "XPML11", TipoProvento.RENDIMENTO_FII, "110.40"),
        _provento(date(2025, 7, 1), "VALE3", TipoProvento.REEMBOLSO_BTC, "35"),
        _provento(date(2026, 3, 10), "ITSA4", jcp, "425", ano=2025),
    )
    f = fichas(_dados(Livro(), 2025, proventos))
    assert f is not None
    por = {(r.codigo, r.cnpj): r.valor for r in f.rendimentos}
    assert por == {("09", PETR.cnpj): D("300"), ("09", ITUB.cnpj): D("100"),
                   ("99", XPML.cnpj): D("110.40"), ("99", VALE.cnpj): D("35"),
                   ("10", ITSA.cnpj): D("425")}  # fmt: skip
    [jcp_bem] = f.bens
    assert (jcp_bem.grupo, jcp_bem.codigo, jcp_bem.cnpj, jcp_bem.anterior, jcp_bem.atual) == (
        "99", "07", ITSA.cnpj, D("0"), D("425"))  # fmt: skip
    g = fichas(_dados(Livro(), 2026, proventos))
    assert g is not None and g.provisorio
    assert [(b.anterior, b.atual) for b in g.bens] == [(D("425"), D("0"))]
    assert not [r for r in g.rendimentos if r.codigo == "10"]  # declarado no ano do crédito


def test_ficha_mes_a_mes_e_imposto_pago_de_dezembro() -> None:
    """Caso B (abril a junho de 2025) e saldo de dezembro: abril V AAAA3 (+1.000) e V CCCC3
    (-2.000), vendas de 13.000: prejuízo de 1.000 compensável (R5), "Prejuízo a compensar"
    1.000,00; maio "Resultado negativo até o mês anterior" 1.000,00. Dezembro: V 4.000 PETR4 a
    30 (resultado 0) com IRRF 6,00 que sobra: ficha "Imposto pago", Lei 11.033 = 6,00."""
    livro = Livro()
    livro.nota(date(2025, 1, 10), livro.op("C", "AAAA3", 100, "40"),
               livro.op("C", "CCCC3", 200, "50"), livro.op("C", "PETR4", 4000, "30"))  # fmt: skip
    livro.nota(date(2025, 4, 10), livro.op("V", "AAAA3", 100, "50"),
               livro.op("V", "CCCC3", 200, "40"))  # fmt: skip
    livro.nota(date(2025, 12, 4), livro.op("V", "PETR4", 4000, "30"), irrf="6.00")
    d = _dados(livro, 2025)
    f = fichas(d)
    assert f is not None and [m.mes for m in f.renda_variavel] == list(range(1, 13))
    abril, maio = f.renda_variavel[3], f.renda_variavel[4]
    assert abril.valor("Mercado à Vista - Ações").comuns == D("-1000.00")
    assert abril.valor(PREJUIZO).comuns == D("1000.00")
    assert maio.valor(NEGATIVO_ANTERIOR).comuns == D("1000.00")
    [pago] = [r for r in f.rendimentos if r.ficha == "imposto_pago"]
    assert (pago.codigo, pago.valor) == ("03", D("6.00"))
    rel = declaracao(d)
    assert [s.titulo for s in rel.secoes][:5] == [
        "Bens e Direitos",
        "Dívidas e ônus reais",
        "Rendimentos isentos e não tributáveis",
        "Rendimentos sujeitos à tributação exclusiva/definitiva",
        "Imposto pago/retido",
    ]
    assert rel.secoes[5].titulo == "Renda variável - operações comuns/day trade - Janeiro"


def test_ano_sem_mapeamento_avisa() -> None:
    d = _dados(Livro(), 2025).model_copy(update={"irpf": None})
    assert fichas(d) is None
    assert "falta data/irpf/2026.yaml" in declaracao(d).notas[0]


def test_opcao_lancada_vai_para_dividas_pelo_premio() -> None:
    """Caso E (parte lançada): V 1.000 VALEA600 a 0,30 em 05/12/2025, vence em 16/01/2026. Não é
    bem (P&R 720 / MIR); por decisão do usuário (20/10/2026, sem norma) vai para Dívidas e ônus
    reais, código 16, pelo prêmio recebido: 0,00 em 31/12/2024 e 300,00 em 31/12/2025; em 2026 a
    posição virou pó e a dívida vai a 0,00."""
    livro = Livro()
    venc = date(2026, 1, 16)
    livro.nota(date(2025, 12, 5), livro.op("V", "VALEA600", 1000, "0.30", mercado="CALL",
                                           vencimento=venc, objeto=TipoAtivo.ACAO))  # fmt: skip
    f = fichas(_dados(livro, 2025))
    assert f is not None and f.bens == ()
    [x] = f.dividas
    assert (x.codigo, x.ticker, x.anterior, x.atual) == ("16", "VALEA600", D("0"), D("300.00"))
    assert x.discriminacao == (
        "Opção de compra VALEA600 lançada (1.000), vencimento 16/01/2026: "
        "prêmio recebido, posição em aberto, corretora XP."
    )
    g = fichas(_dados(livro, 2026))
    assert g is not None and [(y.anterior, y.atual) for y in g.dividas] == [(D("300.00"), D("0"))]
    assert REFS.tipos.tipo("VALE3") is TipoAtivo.ACAO
