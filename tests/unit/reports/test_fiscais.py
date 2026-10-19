"""Relatórios fiscais (F5) a partir de apurações sintéticas: os números de cada relatório saem
da apuração do motor sem cálculo novo (regra 12) e fecham entre si."""

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from coir import referencias
from coir.domain.fiscal import Categoria, Grupo, Visao
from coir.domain.relatorio import DadosRelatorio, NotaDoLivro
from coir.parsing.extrair import extrair_textos
from coir.parsing.normalizar import classes_das_series, operacoes
from coir.reports.fiscais import TODOS, relatorios
from coir.reports.modelo import Relatorio, Secao
from tests.apoio import cpf_de_teste
from tests.fiscal_cases.apoio import Livro, categoria, mes

D = Decimal
FIXTURES = Path(__file__).parents[2] / "fixtures" / "notas"
AGORA = datetime(2026, 10, 16, 9, 0)


def _livro() -> Livro:
    livro = Livro()
    livro.nota(date(2025, 1, 2), livro.op("C", "AAAA3", 1000, "10"), custos="10")
    livro.nota(date(2025, 2, 3), livro.op("V", "AAAA3", 600, "15"))  # vendas 9.000: isento
    livro.nota(date(2025, 3, 3), livro.op("V", "AAAA3", 400, "60"))  # vendas 24.000: tributa
    livro.nota(date(2025, 3, 10), livro.op("C", "BBBB3", 100, "20"), livro.op("V", "BBBB3", 100,
                                                                              "25"),
               irrf_dt="5")  # fmt: skip
    livro.nota(date(2025, 4, 1), livro.op("C", "CCCC3", 200, "30"))
    return livro


def _dados(livro: Livro, notas: tuple[NotaDoLivro, ...] = ()) -> DadosRelatorio:
    fim = date(2025, 12, 31)
    return DadosRelatorio(
        titular=1, nome="Titular de Teste", cpf=cpf_de_teste(), ano=2025, gerado_em=AGORA,
        codigo_darf="6015", apuracao=livro.apurar(fim), ate=fim, data_carteira=fim,
        visao=Visao.TITULAR,
        carteira=tuple(livro.carteira(fim)), carteira_fiscal=tuple(livro.carteira(fim)),
        contas={1: "1234"}, notas=notas,
    )  # fmt: skip


def _secao(rel: Relatorio, titulo: str) -> Secao:
    return next(s for s in rel.secoes if s.titulo == titulo)


def _coluna(s: Secao, nome: str) -> int:
    return [c.nome for c in s.colunas].index(nome)


def test_todos_os_relatorios_saem_com_cabecalho_do_titular() -> None:
    rels = relatorios(_dados(_livro()))
    assert [r.chave for r in rels] == list(TODOS)
    assert all(r.cabecalho.cpf == cpf_de_teste() and r.cabecalho.gerado_em == AGORA for r in rels)
    assert {r.chave for r in rels if r.pdf} == {
        "resumo-fiscal", "darfs", "irpf-mensal", "irrf-e-exclusiva", "operacoes-finalizadas",
        "carteira", "bens-e-direitos", "proventos",
    }  # fmt: skip


def test_resumo_fiscal_igual_a_apuracao() -> None:
    """PM 10,01 (10.010 / 1.000). Março: comum 24.000 - 4.004 = 19.996 (vendas 24.000 > 20.000),
    IR 2.999,40; day trade 500 a 20% = 100, menos o IRRF de 5 -> DARF 3.094,40. Fevereiro: vendas
    de 9.000 isentas (ganho 9.000 - 6.006 = 2.994)."""
    livro = _livro()
    d = _dados(livro)
    rel = TODOS["resumo-fiscal"](d)
    comum = _secao(rel, "Operações comuns")
    marco = next(x for x in comum.linhas if x[0] == "Março")
    assert (
        marco[_coluna(comum, "Resultado")]
        == categoria(d.apuracao, 2025, 3, Categoria.COMUM).resultado
    )
    assert marco[_coluna(comum, "Imposto")] == D("2999.40")
    fev = next(x for x in comum.linhas if x[0] == "Fevereiro")
    assert fev[_coluna(comum, "Isento")] == D("2994.00")
    consol = _secao(rel, "Consolidação do mês")
    linha = next(x for x in consol.linhas if x[0] == "Março")
    assert linha[_coluna(consol, "DARF")] == mes(d.apuracao, 2025, 3).darf == D("3094.40")
    assert consol.total is not None and consol.total[_coluna(consol, "DARF")] == D("3094.40")


def test_darfs_e_ordem_do_irpf_batem_com_a_apuracao() -> None:
    d = _dados(_livro())
    darfs = TODOS["darfs"](d)
    pagar = darfs.secoes[0]
    assert [x[0] for x in pagar.linhas] == ["03/2025"] and pagar.linhas[0][2] == D("3094.40")
    irpf = TODOS["irpf-mensal"](d)
    mercados = _secao(irpf, "Resultados por mercado")
    marco = next(x for x in mercados.linhas if x[0] == "Março")
    m = mes(d.apuracao, 2025, 3)
    dt = next(c for c in m.categorias if c.categoria is Categoria.DAY_TRADE)
    assert marco[_coluna(mercados, "Day trade: À vista")] == dt.por_grupo[Grupo.A_VISTA] == D(500)
    consol = _secao(irpf, "Consolidação do mês")
    linha = next(x for x in consol.linhas if x[0] == "Março")
    assert linha[_coluna(consol, "IR fonte de day trade no mês")] == D(5)
    assert linha[_coluna(consol, "IR fonte (Lei 11.033/2004) no mês")] == 0
    outras = _secao(irpf, "Outras informações (só para consulta)")
    fev = next(x for x in outras.linhas if x[0] == "Fevereiro")
    assert fev[2] == D("2994.00")


def test_operacoes_finalizadas_somam_o_resultado_do_mes() -> None:
    d = _dados(_livro())
    rel = TODOS["operacoes-finalizadas"](d)
    assert [s.titulo for s in rel.secoes] == [
        "2025 - Fevereiro - Bovespa", "2025 - Março - Bovespa"]  # fmt: skip
    marco = rel.secoes[1]
    res = _coluna(marco, "Resultado")
    m = mes(d.apuracao, 2025, 3)
    assert marco.total is not None and marco.total[res] == sum(c.resultado for c in m.categorias)
    obs = {x[1]: x[_coluna(marco, "Obs.")] for x in marco.linhas}
    fev = _secao(rel, "2025 - Fevereiro - Bovespa").linhas[0][-1]
    assert obs["BBBB3"] == "DT" and isinstance(fev, str) and "isento" in fev


def test_conferencia_e_carteira_fecham() -> None:
    d = _dados(_livro())
    conf = TODOS["conferencia-pm"](d)
    cccc = _secao(conf, "CCCC3")
    ultimo = cccc.linhas[-1]
    [pos] = [p for p in d.carteira if p.ticker == "CCCC3"]
    assert ultimo[_coluna(cccc, "Qtd. após")] == pos.quantidade == 200
    assert ultimo[_coluna(cccc, "PM após")] == pos.pm
    carteira = TODOS["carteira"](d)
    acoes = carteira.secoes[0]
    assert [x[0] for x in acoes.linhas] == ["CCCC3"]
    assert acoes.total is not None and acoes.total[-1] == pos.custo
    bens = TODOS["bens-e-direitos"](d)
    assert bens.secoes[0].linhas == (("CCCC3", D(200), pos.pm, pos.custo),)


def test_livro_de_notas_com_custos_por_tipo() -> None:
    texto = (FIXTURES / "rico_2015_11_bovespa.txt").read_text(encoding="utf-8")
    res = extrair_textos("rico.txt", texto.split("\f"))
    refs = referencias.carregar()
    series = classes_das_series(res.notas, refs)
    notas = tuple(NotaDoLivro(nota=n, operacoes=tuple(operacoes(n, refs, series)))
                  for n in res.notas)  # fmt: skip
    rel = TODOS["livro-de-notas"](_dados(_livro(), notas))
    secao = _secao(rel, "Notas")
    [linha] = secao.linhas
    resumo = res.notas[0].resumo_bovespa
    assert resumo is not None
    assert linha[_coluna(secao, "Custos")] == sum(c.valor for c in resumo.custos)
    assert linha[_coluna(secao, "Valor das operações")] == resumo.valor_operacoes
    ops = _secao(rel, "Operações (Bovespa)")
    assert len(ops.linhas) == len(notas[0].operacoes)


def test_ordem_do_irpf_separa_a_ficha_de_fii() -> None:
    """FII tem ficha própria no programa IRPF: o total de comuns/day trade não leva o imposto de
    FII. Junho/2024: comum 30.000 - 20.000 = 10.000 (IR 1.500); FII 11.000 - 10.000 = 1.000 a 20%
    (IR 200). IRRF de 2 repartido pelas vendas (FII 2 x 11.000/41.000 = 0,54; comuns 1,46): cada
    ficha abate o seu -> 1.498,54 + 199,46 = DARF 1.698,00 (a do saldo único)."""
    from coir.domain.fiscal import TipoAtivo

    livro = Livro()
    livro.nota(date(2024, 6, 3), livro.op("C", "AAAA3", 2000, "10"),
               livro.op("C", "XPML11", 100, "100", tipo=TipoAtivo.FII))  # fmt: skip
    livro.nota(date(2024, 6, 10), livro.op("V", "AAAA3", 2000, "15"),
               livro.op("V", "XPML11", 100, "110", tipo=TipoAtivo.FII), irrf="2")  # fmt: skip
    fim = date(2024, 12, 31)
    d = _dados(livro).model_copy(update={"ano": 2024, "apuracao": livro.apurar(fim), "ate": fim})
    rel = TODOS["irpf-mensal"](d)
    consol = _secao(rel, "Consolidação do mês")
    junho = next(x for x in consol.linhas if x[0] == "Junho")
    assert junho[_coluna(consol, "Total do imposto devido")] == D("1500.00")
    assert junho[_coluna(consol, "IR fonte (Lei 11.033/2004) no mês")] == D("1.46")
    assert junho[_coluna(consol, "Imposto a pagar")] == D("1498.54")
    assert junho[_coluna(consol, "Imposto pago (calculado)")] == D("1498.54")
    fii = _secao(rel, "Fundos imobiliários e Fiagro")
    jfii = next(x for x in fii.linhas if x[0] == "Junho")
    assert jfii[_coluna(fii, "IR fonte (Lei 11.033/2004) no mês")] == D("0.54")
    assert jfii[_coluna(fii, "Imposto a pagar")] == D("199.46")
    darf = next(m.darf for m in d.apuracao.meses if m.mes == 6)
    assert darf == D("1698.00") == D("1498.54") + D("199.46")
