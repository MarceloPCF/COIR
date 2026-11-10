"""FI-Infra (Lei 12.431/2011 art. 3º) e FIP-IE (Lei 11.478/2007 art. 2º §1º I), tipo FUNDO_INFRA
pela lista de fundos da B3: o ganho de pessoa física tem alíquota zero, fica fora da apuração
mensal, do limite de R$ 20 mil e do DARF, e o prejuízo não compensa outros ganhos (IN RFB
1.585/2015 art. 64 só admite as perdas dos arts. 27, 58 e 60 a 62). Parecer do especialista de
09/11/2026, seguido pelo usuário."""

from datetime import date

from coir.domain.fiscal import Categoria, TipoAtivo
from tests.fiscal_cases.apoio import D, Livro, categoria

INFRA = TipoAtivo.FUNDO_INFRA


def test_ganho_de_fi_infra_fica_fora_da_apuracao_e_do_limite() -> None:
    """Fevereiro: ações 19.000 (isentas) + JURO11 11.000; sem o fundo na soma, as ações seguem
    isentas, e o ganho de 1.000 do fundo vai para a lista à parte, sem imposto."""
    livro = Livro()
    livro.nota(date(2025, 1, 10), livro.op("C", "PETR4", 1000, "10"),
               livro.op("C", "JURO11", 100, "100", tipo=INFRA))  # fmt: skip
    livro.nota(date(2025, 2, 10), livro.op("V", "PETR4", 1000, "19"),
               livro.op("V", "JURO11", 100, "110", tipo=INFRA))  # fmt: skip
    ap = livro.apurar(date(2025, 2, 28))
    fevereiro = categoria(ap, 2025, 2, Categoria.COMUM)
    assert (fevereiro.vendas_acoes, fevereiro.isento, fevereiro.imposto) == (D(19000), D(9000), 0)
    [fundo] = ap.exclusiva
    assert (fundo.ticker, fundo.resultado, fundo.tipo) == ("JURO11", D(1000), INFRA)


def test_prejuizo_de_fip_ie_nao_compensa_ganho_de_acoes() -> None:
    """Prejuízo de 5.000 no FIP-IE em março não reduz o ganho tributável de 10.000 em abril."""
    livro = Livro()
    livro.nota(date(2025, 3, 3), livro.op("C", "XPIE11", 100, "100", tipo=INFRA),
               livro.op("C", "PETR4", 1000, "20"))  # fmt: skip
    livro.nota(date(2025, 3, 20), livro.op("V", "XPIE11", 100, "50", tipo=INFRA))
    livro.nota(date(2025, 4, 10), livro.op("V", "PETR4", 1000, "30"))
    ap = livro.apurar(date(2025, 4, 30))
    assert categoria(ap, 2025, 3, Categoria.COMUM).prejuizo_a_compensar == 0
    assert categoria(ap, 2025, 4, Categoria.COMUM).imposto == D("1500.00")
    assert [e.resultado for e in ap.exclusiva] == [D(-5000)]


def test_day_trade_de_fi_infra_tambem_fica_fora() -> None:
    livro = Livro()
    livro.nota(date(2025, 5, 5), livro.op("C", "JURO11", 100, "100", tipo=INFRA),
               livro.op("V", "JURO11", 100, "101", tipo=INFRA))  # fmt: skip
    ap = livro.apurar(date(2025, 5, 31))
    assert categoria(ap, 2025, 5, Categoria.DAY_TRADE).imposto == 0
    assert [e.resultado for e in ap.exclusiva] == [D(100)]


def test_bens_e_direitos_e_ficha_do_ganho_na_dirpf() -> None:
    """Parecer de 09/11/2026 (Manual MIR; Guia IR 2026): FI-Infra em 07.10 e FIP-IE em 07.07, com
    o CNPJ do fundo; o ganho na ficha Exclusiva/Definitiva, código 06 (Lei 12.431 art. 3º §2º:
    "exclusivamente na fonte"), não em isentos."""
    from coir.domain.fiscal import PosicaoAtivo
    from coir.domain.relatorio import InfoAtivo
    from coir.referencias import carregar_irpf
    from coir.reports.irpf import _tipo_bem  # pyright: ignore[reportPrivateUsage]

    pos = PosicaoAtivo(titular=1, corretora=None, conta=None, ativo=1, ticker="XPIE11", tipo=INFRA,
                       vencimento=None, quantidade=D(10), custo=D(1000))  # fmt: skip

    def info(tipo_b3: str) -> InfoAtivo:
        return InfoAtivo(emissor="XPIE", cnpj="1", razao_social="X", especificacao="CI",
                         tipo_b3=tipo_b3)  # fmt: skip

    mapa = carregar_irpf(2025)
    assert mapa.bens.item(_tipo_bem(pos, info("FIP-IE")) or "ACAO").codigo == "07"
    assert mapa.bens.item(_tipo_bem(pos, info("FI-INFRA")) or "ACAO").codigo == "10"
    assert mapa.exclusivos.linha("ganho_fundo_infra").codigo == "06"


def test_rendimentos_distribuidos_vao_para_a_exclusiva_06() -> None:
    """Parecer de 09/11/2026: rendimento de FI-Infra (Lei 12.431 art. 3º §§1º e 2º) e de FIP-IE
    (Lei 11.478 art. 2º §3º diz isento, mas o programa só aceita a exclusiva 06) na exclusiva 06,
    em linhas separadas; nunca em isentos de FII."""
    from types import SimpleNamespace

    from coir.domain.relatorio import InfoAtivo
    from coir.referencias import carregar_irpf
    from coir.reports.irpf import _rendimentos_fundo_infra  # pyright: ignore[reportPrivateUsage]

    def info(tipo_b3: str) -> InfoAtivo:
        return InfoAtivo(emissor="X", cnpj="1", razao_social="X", especificacao="CI",
                         tipo_b3=tipo_b3)  # fmt: skip

    d = SimpleNamespace(info_ativos={"XINF11": info("FI-INFRA"), "XPIE11": info("FIP-IE")})
    feitas: list[tuple[str, str, list[tuple[str, D]]]] = []
    mapa = carregar_irpf(2026)
    _rendimentos_fundo_infra(d, [("XINF11", D(10)), ("XPIE11", D(20))],  # type: ignore[arg-type]
                             lambda f, c, i: feitas.append((f, c, list(i))), mapa)  # fmt: skip
    assert feitas == [("exclusivos", "rendimento_fundo_infra", [("XINF11", D(10))]),
                      ("exclusivos", "rendimento_fip_ie", [("XPIE11", D(20))])]  # fmt: skip
    assert mapa.exclusivos.linha("rendimento_fip_ie").codigo == "06"
    assert not mapa.isentos.tem("rendimento_fip_ie")


def test_amortizacao_acima_do_custo_zera_o_custo_e_avisa_o_excedente() -> None:
    """Decisão do usuário de 10/11/2026 (leitura conservadora): C 10 JURO11 a 100,00 (custo
    1.000,00); amortização de 1.200,00 em 15/06/2025. O custo vai a zero e o excedente de 200,00 é
    rendimento do fundo (Lei 11.478/2007 art. 2º §2º; Lei 12.431/2011 art. 3º), com alíquota zero:
    o COIR avisa e não o lança. Antes, o motor falhava."""
    from coir.domain.fiscal import TipoEvento

    livro = Livro()
    livro.nota(date(2025, 1, 10), livro.op("C", "JURO11", 10, "100", tipo=INFRA))
    livro.evento(date(2025, 6, 15), "JURO11", TipoEvento.REDUCAO_CAPITAL, "V", None, "1200",
                 tipo_ativo=INFRA)  # fmt: skip
    [posicao] = livro.carteira(date(2025, 6, 30))
    assert (posicao.quantidade, posicao.custo) == (D(10), 0)
    ap = livro.apurar(date(2025, 6, 30))
    [aviso] = [a for a in ap.avisos if "amortização" in a]
    assert aviso.startswith("15/06/2025: amortização de JURO11 acima do custo")
    assert "R$ 200.00" in aviso
