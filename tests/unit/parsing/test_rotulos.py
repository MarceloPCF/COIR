"""Rótulos das notas (coir/parsing/rotulos.yaml): as variações aceitas de cada um e, tão importante
quanto, as que não podem casar com outro campo da nota."""

import re
import time
from decimal import Decimal
from pathlib import Path

import pytest

from coir.domain.nota import Montante
from coir.erros import ExtracaoError
from coir.parsing import rotulos
from coir.parsing.extrair import extrair_textos
from coir.parsing.rotulos import linha_inteira, no_inicio, procurar, valor_apos

PASTA = Path(__file__).parents[2] / "fixtures" / "notas"
FONTES = Path(__file__).parents[3] / "coir" / "parsing"


def _casa(chave: str, texto: str) -> bool:
    return rotulos.regex(chave).search(texto) is not None


# ------------------------------------------------------------------------- variações aceitas
ACEITAS: dict[str, tuple[str, ...]] = {
    "numero_nota": ("Nr. nota", "Nr.Nota", "NR NOTA", "Nr nota", "Nº nota", "N° da nota",
                    "No. nota", "No nota", "Nro. nota", "Num. nota", "Número da nota",
                    "NUMERO DA NOTA", "Nr.   nota"),
    "folha": ("Folha", "FOLHA", "Fl.", "Fl 1", "Fls."),
    "data_pregao": ("Data pregão", "Data pregao", "DATA PREGÃO", "Data do pregão", "Data de pregão",
                    "Datapregão", "Data   pregão", "D. pregão", "D.pregão",
                    "Dt. pregão", "Dt pregão", "Dt. do pregão", "data do pregao"),
    "nota_de_corretagem": ("NOTA DE CORRETAGEM", "Nota da corretagem", "Notadecorretagem"),
    "codigo_cliente": ("Código cliente", "Codigo do Cliente", "Cód. cliente", "Cod cliente",
                       "CÓDIGO DO CLIENTE"),
    "cv_mercadoria": ("C/V Mercadoria", "C/VMercadoria"),
    "resumo_financeiro": ("Resumo Financeiro", "RESUMOFINANCEIRO"),
    "corretagem_despesas": ("Corretagem / Despesas", "Corretagem e despesas"),
    "day_trade": ("Day Trade", "Day-Trade", "DAYTRADE", "day - trade"),
    "resumo_negocios": ("Resumo dos Negócios", "Resumo de negocios", "RESUMODOSNEGOCIOS"),
    "vendas_a_vista": ("Vendas à vista", "Vendas a vista", "Vendas vista", "VENDASAVISTA"),
    "compras_a_vista": ("Compras à vista", "Compras a vista"),
    "opcoes_compras": ("Opções - compras", "Opções compras", "OPCOES-COMPRAS"),
    "opcoes_vendas": ("Opções - vendas", "Opcoes vendas"),
    "operacoes_a_termo": ("Operações à termo", "Operações a termo", "Operacoes termo"),
    "valor_operacoes_titulos_publicos": ("Valor das oper. c/ títulos públ.",
                                         "Valor das operações com títulos públicos"),
    "valor_operacoes": ("Valor das operações", "Valor das operacoes", "Vlr. das operações",
                        "Valor de operações", "Valordasoperações"),
    "valor_liquido_operacoes": ("Valor líquido das operações", "Valor liq. das operações",
                                "Vlr. líquido das operações", "Valor líquido operações"),
    "taxa_liquidacao": ("Taxa de liquidação", "Taxa de liquidação/CCP", "Taxa deliquidação",
                        "Taxa liquidação", "Tx. liquidação", "Tx de liquidação"),
    "taxa_registro": ("Taxa de Registro", "Taxa registro", "Tx. de registro"),
    "taxa_termo_opcoes": ("Taxa de termo/opções", "Taxa termo e opções", "Tx. termo/opções"),
    "taxa_ana": ("Taxa A.N.A.", "Taxa ANA", "Tx. A.N.A"),
    "emolumentos": ("Emolumentos", "EMOLUMENTOS"),
    "taxa_transferencia_ativos": ("Taxa de Transf. de Ativos", "Taxa de Transferencia de Ativos",
                                  "Taxa de Transferência de Ativos", "Tx. transf. ativos"),
    "corretagem": ("Corretagem", "CORRETAGEM"),
    "taxa_operacional": ("Taxa Operacional", "TaxaOperacional", "Tx. operacional"),
    "execucao_casa": ("Execução casa", "Execucao casa"),
    "taxa_custodia": ("Taxa de Custódia", "Taxa custodia", "Tx. de custódia"),
    "iss": ("ISS (SÃO PAULO)", "I.S.S", "I.S.S."),
    "total_cblc": ("Total CBLC", "Total da CBLC", "TotalCBLC"),
    "total_bovespa": ("Total Bovespa / Soma", "Total Bovespa", "Total Bovespa Soma"),
    "total_despesas": ("Total Custos / Despesas", "Total corretagem / Despesas",
                       "Total custos e despesas", "Total das despesas", "Total despesas"),
    "irrf_operacoes": ("I.R.R.F. s/ operações", "IRRF s/ operações", "I.R.R.F. sobre operações",
                       "IRRF s/ opera��es",  # acento ilegível no PDF
                       "IRRFs/operações"),
    "liquido_para": ("Líquido para", "Liquido para", "Líq. para", "LIQUIDO PARA"),
    "observacoes": ("(*) - Observações", "(*) Observações", "(*) - Obs."),
    "venda_disponivel": ("Venda disponível", "Venda disp.", "Vendadisponivel"),
    "irrf_day_trade": ("IRRF Day Trade", "IRRF Day-Trade", "IRRFDay Trade", "I.R.R.F. Day Trade"),
    "outros_custos": ("+Outros Custos", "+ Outros custos", "Outros custos"),
    "irrf_operacional": ("IRRF operacional", "I.R.R.F. Operacional"),
}  # fmt: skip


@pytest.mark.parametrize(
    ("chave", "texto"), [(c, t) for c, textos in ACEITAS.items() for t in textos]
)
def test_variacao_aceita(chave: str, texto: str) -> None:
    assert _casa(chave, texto), f"{chave} deveria aceitar {texto!r}"


def test_todo_rotulo_do_yaml_tem_variacoes_testadas_ou_e_palavra_unica() -> None:
    simples = {c for c in rotulos.chaves() if all(" " not in f for f in rotulos.formas(c))}
    assert set(rotulos.chaves()) - simples <= set(ACEITAS)


def test_todo_rotulo_do_yaml_e_usado_no_parser() -> None:
    codigo = "\n".join(p.read_text(encoding="utf-8") for p in FONTES.glob("*.py"))
    sem_uso = [c for c in rotulos.chaves() if f'"{c}"' not in codigo and f"'{c}'" not in codigo]
    assert sem_uso == []


# ------------------------------------------------------------- confusões que não podem ocorrer
RECUSADAS: tuple[tuple[str, str], ...] = (
    # "Data" sozinha e outras datas da nota
    ("data_pregao", "Data liquidação 05/01/2024"),
    ("data_pregao", "Data de liquidação"),
    ("data_pregao", "Data"),
    ("data_pregao", "Pregão"),
    ("data_pregao", "Dtpregão"),  # abreviação sem ponto não pode grudar na palavra seguinte
    ("data_pregao", "Validade pregão"),  # não começa no meio de uma palavra
    ("data_pregao", "D pregão"),  # letra solta só com o ponto: "D" pode ser o D/C de um valor
    ("data_pregao", "12,00 D pregão"),
    # número da nota x título da nota e endereço
    ("numero_nota", "NOTA DE CORRETAGEM"),
    ("numero_nota", "NOTA DE NEGOCIAÇÃO"),
    ("numero_nota", "Av. Fictícia, nº 100 - 2º ANDAR"),
    ("numero_nota", "Nonota"),
    ("nota_de_corretagem", "Nr. nota"),
    # folha x rodapé e palavras que começam por "Fl"
    ("folha", "Página 1 de 2"),
    ("folha", "Pág. 1"),
    ("folha", "Florianópolis"),
    ("folha", "Flex"),
    # taxas entre si e com a data de liquidação
    ("taxa_liquidacao", "Data liquidação"),
    ("taxa_liquidacao", "Taxa de Registro"),
    ("taxa_registro", "Taxa de liquidação"),
    ("taxa_operacional", "Corretagem"),
    ("corretagem", "Taxa Operacional"),
    ("corretagem", "Clearing"),
    ("clearing", "Corretagem"),
    ("taxa_ana", "Taxa de termo/opções"),
    ("taxa_ana", "Taxa Analítica"),
    ("taxa_termo_opcoes", "Operações à termo"),
    ("taxa_custodia", "Taxa de liquidação"),
    # emolumentos x taxas BM&F (emol+f.gar): nunca entram os dois
    ("emolumentos", "Taxas BM&F (emol+f.gar)"),
    ("emolumentos", "Taxa registro BM&F + (emol+f.gar)"),
    # valores do resumo entre si
    ("valor_operacoes", "Valor líquido das operações"),
    ("valor_operacoes", "Valor das oper. c/ títulos públ."),
    ("valor_operacoes", "Valor dos negócios"),
    ("valor_operacoes", "Valor"),
    ("valor_liquido_operacoes", "Valor das operações"),
    ("valor_operacoes_titulos_publicos", "Valor das operações"),
    ("vendas_a_vista", "Compras à vista"),
    ("opcoes_compras", "Opções - vendas"),
    ("operacoes_a_termo", "Taxa de termo/opções"),
    # totais entre si
    ("total_cblc", "Total Bovespa / Soma"),
    ("total_bovespa", "Total CBLC"),
    ("total_despesas", "Total CBLC"),
    ("total_despesas", "Total corretagem"),
    ("total", "Total CBLC"),
    # IRRF sobre operações x IRRF day trade
    ("irrf_operacoes", "IRRF Day Trade: Base R$ 0,00 Projeção R$ 0,00"),
    ("irrf_day_trade", "I.R.R.F. s/ operações, base R$0,00"),
    ("irrf_operacional", "IRRF Day Trade (proj.)"),
    ("liquido_para", "Valor líquido das operações"),
    ("execucao_casa", "Execução"),
    ("iss", "Comissão"),  # "iss" no meio da palavra
    ("cv_mercadoria", "C/V Tipo mercado"),
    ("codigo_cliente", "Cliente"),
    ("outros_custos", "Outros IRRF Operacional"),
)


@pytest.mark.parametrize(("chave", "texto"), RECUSADAS)
def test_variacao_nao_casa_outro_campo(chave: str, texto: str) -> None:
    alvo = rotulos.regex(chave)
    casou = alvo.fullmatch(texto) if chave == "total" else alvo.search(texto)
    assert casou is None, f"{chave} não pode casar {texto!r}"


def test_iss_le_so_o_proprio_rotulo() -> None:
    # "ISSQN" começa como "ISS": a forma sem ponto aceita, mas o valor tem de vir logo depois
    assert valor_apos("ISSQN 1,00", "iss") is None


def test_rotulo_parecido_sem_valor_logo_depois_nao_da_valor() -> None:
    """Prefixo de outro rótulo: "Execução" está em "Execução casa", "Taxa registro" em "Taxa
    registro BM&F"; o valor só vale se vier logo depois do rótulo inteiro."""
    assert valor_apos("Execução casa 1,00", "execucao") is None
    assert valor_apos("Valor das operações com títulos públicos 1,00", "valor_operacoes") is None
    assert valor_apos("Emolumentos/Fees 1,00", "emolumentos") is None
    assert valor_apos("Taxa registro BM&F 0,50", "taxa_registro") is None
    assert valor_apos("Corretagem / Despesas\nCorretagem 2,00 D", "corretagem") == Montante(
        valor=Decimal(2), dc="D"
    )


def test_valor_maiusculo_depois_do_rotulo() -> None:
    """Só o rótulo ignora a caixa: o D/C que vem depois continua em maiúscula."""
    assert valor_apos("taxa de liquidação 1,49 d", "taxa_liquidacao") == Montante(
        valor=Decimal("1.49")
    )


# ---------------------------------------------------------------- sintaxe e busca nas linhas
def test_abreviacao_e_ordinal() -> None:
    assert _casa("numero_nota", "N° nota") and _casa("numero_nota", "Nº nota")
    assert no_inicio("  Fl.1", "folha") and not no_inicio("Flexível", "folha")


def test_rotulo_colado_so_quando_pedido() -> None:
    assert not _casa("day_trade", "IRRFDay Trade")
    assert re.search(rotulos.padrao("day_trade", colado=True), "IRRFDay Trade")


def test_sequencia_de_rotulos_na_mesma_linha() -> None:
    assert linha_inteira("Nr. Nota Data pregão", "numero_nota", "data_pregao")
    assert linha_inteira("Nº NotaDatapregão", "numero_nota", "data_pregao")
    assert not linha_inteira("Nr.Nota Folha Data pregão", "numero_nota", "data_pregao")
    assert linha_inteira("Nr.Nota Folha Data pregão", "numero_nota", "folha", "data_pregao")


def test_rotulo_quebrado_entre_duas_linhas() -> None:
    linhas = ["Data", "pregão", "02/01/2024"]
    assert procurar(linhas, "data_pregao") == (0, 1)


def test_linha_inteira_vence_a_quebrada() -> None:
    linhas = ["Data", "pregão", "x", "Data pregão", "02/01/2024"]
    assert procurar(linhas, "data_pregao") == (3, 3)


def test_rotulo_ausente_falha_alto_com_a_forma_de_referencia() -> None:
    with pytest.raises(ExtracaoError, match="Data do pregão"):
        procurar(["Data liquidação", "05/01/2024"], "data_pregao")


def test_forma_de_referencia_casa_o_proprio_rotulo() -> None:
    for chave in rotulos.chaves():
        assert _casa(chave, rotulos.referencia(chave)), chave
    assert rotulos.referencia("taxa_liquidacao") == "Taxa de liquidação/CCP"


def test_forma_so_com_palavras_opcionais_e_erro_de_configuracao(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from coir.erros import ConfiguracaoError

    monkeypatch.setattr(rotulos, "formas", lambda chave: ("[de|do]",))
    with pytest.raises(ConfiguracaoError, match="forma vazia"):
        rotulos.padrao.__wrapped__("qualquer")


def test_pontos_opcionais_sem_retrocesso_explosivo() -> None:
    """Linha de PDF com colunas cheias de espaço: a regex responde na hora, case ou não."""
    linha = "I" + " " * 300 + "R" + " " * 300 + "R" + " " * 300 + "X"
    inicio = time.perf_counter()
    assert rotulos.regex("irrf_operacoes").search(linha) is None
    assert rotulos.regex("taxa_ana").search("Taxa A" + " " * 300 + "N" + " " * 300 + "X") is None
    assert time.perf_counter() - inicio < 1


def test_guardas_veem_o_rotulo_mesmo_colado() -> None:
    """Guardas e âncoras não usam a fronteira à esquerda: rótulo colado ao texto de antes ainda
    faz a nota falhar alto ou ser lida no segmento certo, em vez de sumir calado."""
    from coir.domain.nota import Segmento
    from coir.parsing import bovespa
    from coir.parsing.extrair import _conta_bmf
    from coir.parsing.paginas import pagina_generica

    with pytest.raises(ExtracaoError, match="fora do padrão"):
        bovespa._irrf("XI.R.R.F. s/ operações, base R$ 1,00", "xp")
    texto = "Endereço SPCódigo cliente\nRua Fictícia, 10 4567\nRodapé 0800"
    assert _conta_bmf(texto, "btg") == "4567"
    pagina = pagina_generica("DescriçãoC/V Mercadoria Vencimento", "xp", "CORRETORA 1")
    assert pagina.segmento is Segmento.BMF


def test_rotulo_desconhecido_e_erro_de_configuracao() -> None:
    from coir.erros import ConfiguracaoError

    with pytest.raises(ConfiguracaoError):
        rotulos.regex("nao_existe")


# ------------------------------------------- a nota inteira com os rótulos escritos de outro jeito
# (forma impressa nas notas fictícias, variação que outra corretora poderia imprimir)
OUTRAS_GRAFIAS: tuple[tuple[str, str], ...] = (
    ("Nr. nota", "Nº da nota"),
    ("Nr.Nota", "Número da Nota"),
    ("Data pregão", "DATA DO PREGAO"),
    ("Código cliente", "Cód. do cliente"),
    ("Codigo do Cliente", "Cód. Cliente"),
    ("Resumo dos Negócios", "Resumo de negocios"),
    ("Vendas à vista", "Vendas a vista"),
    ("Opções - compras", "Opções compras"),
    ("Valor líquido das operações", "Vlr. líq. das operações"),
    ("Valor das operações", "Valor das  operacoes"),
    ("Taxa de liquidação", "Tx. liquidação"),
    ("Taxa de Registro", "Taxa registro"),
    ("Taxa A.N.A.", "Taxa ANA"),
    ("Taxa de termo/opções", "Tx. termo e opções"),
    ("Total CBLC", "Total da CBLC"),
    ("Total Bovespa / Soma", "TOTAL BOVESPA"),
    ("I.R.R.F. s/ operações", "IRRF sobre operações"),
    ("Líquido para", "Líq. para"),
    ("Venda disponível", "Venda disp."),
)
FIXTURES = sorted(p.name.removesuffix(".txt") for p in PASTA.glob("*.txt"))


@pytest.mark.parametrize("nome", FIXTURES)
def test_nota_com_outras_grafias_extrai_igual(nome: str) -> None:
    texto = (PASTA / f"{nome}.txt").read_text(encoding="utf-8")
    trocado = texto
    for impresso, variacao in OUTRAS_GRAFIAS:
        trocado = trocado.replace(impresso, variacao)
    assert trocado != texto
    original = extrair_textos(f"{nome}.pdf", texto.split("\f"))
    variado = extrair_textos(f"{nome}.pdf", trocado.split("\f"))
    assert variado.rejeitadas == []
    assert variado.notas == original.notas
    assert variado.clientes == original.clientes


def test_cabecalho_quebrado_em_duas_linhas_extrai_igual() -> None:
    texto = (PASTA / "xp_2026_08_bovespa.txt").read_text(encoding="utf-8")
    assert texto.count("Data pregão") == 1
    original = extrair_textos("x.pdf", texto.split("\f"))
    quebrado = extrair_textos("x.pdf", texto.replace("Data pregão", "Data\npregão").split("\f"))
    assert quebrado.rejeitadas == []
    assert quebrado.notas == original.notas


def test_layout_b3rv_com_rotulos_colados_e_abreviados() -> None:
    from tests.unit.parsing.test_layout_b3rv import NOTA

    original = extrair_textos("b3rv.pdf", [NOTA])
    trocado = NOTA.replace("Nr. Nota Data pregão", "Nº Nota Datapregão").replace(
        "Código cliente", "Cód. cliente"
    )
    assert trocado != NOTA
    variado = extrair_textos("b3rv.pdf", [trocado])
    assert variado.rejeitadas == []
    assert variado.notas == original.notas
