"""F8: comparação das fichas do COIR com a declaração entregue (dados fictícios)."""

from decimal import Decimal
from pathlib import Path

from coir import referencias
from coir.domain.irpf import BemIrpf, FichasIrpf, LinhaFicha, MesFichaIrpf, RendimentoIrpf
from coir.gabarito.dirpf import ler_texto
from coir.gabarito.dirpf_comparar import carregar_explicacoes, comparar, explicar, gravar
from tests.unit.test_dirpf_leitura import LINHAS_MERCADO, _paginas

D = Decimal
MAPA = referencias.carregar_irpf(2025)
VALE, ITSA = "33592510000154", "61532644000115"


def _mes(mes: int, acoes: str) -> MesFichaIrpf:
    linhas = [LinhaFicha(rotulo=r, comuns=D(acoes) if i == 0 else D(0),
                         day_trade=D("-17.20") if i == 8 else D(0))
              for i, r in enumerate(LINHAS_MERCADO)]  # fmt: skip
    return MesFichaIrpf(mes=mes, linhas=tuple(linhas))


def _fichas(**troca: object) -> FichasIrpf:
    base: dict[str, object] = dict(
        exercicio=2026, ano=2025, provisorio=False,
        bens=(BemIrpf(tipo="ACAO", grupo="03", codigo="01", ticker="VALE3", cnpj=VALE,
                      discriminacao="", anterior=D("6000"), atual=D("0")),
              BemIrpf(tipo="OPCAO", grupo="04", codigo="04", ticker="ITSAA100", cnpj=None,
                      discriminacao="", anterior=D("0"), atual=D("24"))),
        rendimentos=(
            RendimentoIrpf(ficha="isentos", chave="dividendos", codigo="09", cnpj=VALE,
                           fonte="VALE", valor=D("1000")),
            RendimentoIrpf(ficha="isentos", chave="dividendos", codigo="09", cnpj=ITSA,
                           fonte="ITAUSA", valor=D("300.50")),
            RendimentoIrpf(ficha="isentos", chave="reembolso_btc", codigo="99", cnpj=VALE,
                           fonte="VALE", valor=D("12.34")),
            RendimentoIrpf(ficha="exclusivos", chave="jcp", codigo="10", cnpj=ITSA,
                           fonte="ITAUSA", valor=D("85")),
            RendimentoIrpf(ficha="imposto_pago", chave="ir_fonte_lei_11033", codigo="03",
                           cnpj=None, fonte="", valor=D("4.56")),
        ),
        renda_variavel=tuple(_mes(m, "-1234.56" if m == 1 else "100") for m in range(1, 13)),
        fii=(),
    )  # fmt: skip
    return FichasIrpf.model_validate(base | troca)


def test_tudo_confere() -> None:
    linhas = comparar(_fichas(), ler_texto(_paginas()), MAPA, D("1"), D("0.01"))
    assert {x.classe for x in linhas} == {"confere"}
    assert len([x for x in linhas if x.secao == "bens"]) == 4


def test_divergencias_classificadas_e_explicadas(tmp_path: Path) -> None:
    f = _fichas(bens=(BemIrpf(tipo="ACAO", grupo="03", codigo="01", ticker="VALE3", cnpj=VALE,
                              discriminacao="", anterior=D("6000.02"), atual=D("0")),
                      BemIrpf(tipo="ACAO", grupo="03", codigo="01", ticker="PETR4", cnpj=None,
                              discriminacao="", anterior=D("0"), atual=D("10"))))  # fmt: skip
    linhas = comparar(f, ler_texto(_paginas()), MAPA, D("1"), D("0.01"))
    abertas = {(x.secao, x.item, x.coluna): x.classe for x in linhas if x.classe != "confere"}
    assert abertas == {
        ("bens", "VALE3", "anterior"): "divergente",
        ("bens", "PETR4", "anterior"): "so_coir", ("bens", "PETR4", "atual"): "so_coir",
        ("bens", "ITSAA100", "anterior"): "so_dirpf", ("bens", "ITSAA100", "atual"): "so_dirpf",
    }  # fmt: skip
    arquivo = tmp_path / "dirpf-explicadas.csv"
    arquivo.write_text("secao,item,coluna,motivo\nbens,VALE3,anterior,arredondamento do PM\n",
                       encoding="utf-8")  # fmt: skip
    explicadas = explicar(linhas, carregar_explicacoes(arquivo))
    [vale] = [x for x in explicadas if (x.item, x.coluna) == ("VALE3", "anterior")]
    assert (vale.classe, vale.explicacao, vale.diferenca) == (
        "explicada", "arredondamento do PM", D("0.02"))  # fmt: skip
    saida = tmp_path / "div.csv"
    gravar(explicadas, saida)
    assert saida.read_text(encoding="utf-8").startswith("secao;item;coluna;coir;dirpf")


def test_codigo_antigo_e_fonte_fora_da_bolsa() -> None:
    """O código antigo da DIRPF vai para o atual; numa linha parcial (99) a fonte que o COIR não
    conhece (aplicação, outros) não é divergência."""
    vale5 = BemIrpf(tipo="ACAO", grupo="03", codigo="01", ticker="VALE5", cnpj=VALE,
                    discriminacao="", anterior=D("6000"), atual=D("0"))  # fmt: skip
    f = _fichas(bens=(vale5, _fichas().bens[1]))
    linhas = comparar(f, ler_texto(_paginas()), MAPA, D("1"), D("0.01"),
                      lambda t: "VALE5" if t == "VALE3" else t)  # fmt: skip
    assert {x.classe for x in linhas if x.secao == "bens"} == {"confere"}
    paginas = _paginas()
    paginas[0] = paginas[0].replace(
        "12,34\nTOTAL", "12,34\nTitular 000.000.000-00 00.000.000/0001-91 BANCO 50,00\nTOTAL"
    )
    linhas = comparar(_fichas(), ler_texto(paginas), MAPA, D("1"), D("0.01"))
    assert {x.classe for x in linhas if x.secao == "rendimentos"} == {"confere"}
