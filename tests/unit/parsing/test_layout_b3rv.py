"""Layout b3rv (notas da Clear de 2018 a 2020): rótulos "Nr. Nota Data pregão" colados, negócios
"B3 RV LISTADO" com o C/V colado e o resumo em blocos de rótulos, valores e D/C. Nota inteiramente
fictícia, na mesma ordem de linhas que o PDF dá."""

from datetime import date
from decimal import Decimal

import pytest

from coir.domain.nota import ROTULO_AJUSTE, CategoriaCusto, NotaExtraida, ResumoBovespa
from coir.gabarito.conferencia import linhas_bovespa
from coir.parsing.anonimizar import MapaAnonimizacao, conferir
from coir.parsing.cliente import cliente
from coir.parsing.extrair import extrair_textos
from coir.parsing.paginas import classificar

NOTA = """\
273207
Nr. Nota Data pregão
02/04/2018
Folha
CLEAR CORRETORA - GRUPO XP
Av. Fictícia, 100 - 2º ANDAR - VILA FALSA - SAO PAULO - SP
Tel. (11) 4000-0000
04500-000 SAO PAULO - SP
Internet: www.clear.com.br SAC: 0800.000.0000
C.N.P.J.: 02.332.886/0011-78
Ouvidoria: Tel. 0800-000--0000
e-mail: www.clear.com.br/site/faq/clear-atendimento
Carta Patente:
e-mail ouvidoria: www.clear.com.br/site/ouvidoria
1
 NOTA DE CORRETAGEM
Cliente
765432 - 1 FULANO DE TAL SILVA
C.P.F./C.N.P.J./C.V.M./C.O.B.
000.000.000-00
RUA IMAGINARIA, 45 - BAIRRO INVENTADO 2
13000-000 CIDADEFALSA - SP
Código cliente Assessor
308 - 9 765432 - 1 5
Participante destino do repasse Cliente Valor Custodiante C.I.
0,00 C N
102
Banco Agência
98765
Conta corrente
4567890
Acionista Administrador Qualificado Inv. Qual.
S
Negócios realizados
Q Negociação C/V Tipo mercado Prazo Especificação do título Obs. (*) Quantidade Preço / Ajuste Valor Operação / Ajuste D/C
(019) 99888-7766
B3 RV LISTADOC VISTA 01/00 PETROBRAS PN N2 100 30,000000 3.000,00 D
B3 RV LISTADOV FRACIONARIO 01/00 VALE ON NM # 10 50,500000 505,00 C
Resumo Financeiro D/C
Clearing
Valor líquido das operações 2.495,00
Taxa de liquidação/CCP
Taxa de Registro
Total
0,87
0,00
2.495,87
D
D
D
D
Bolsa
Taxa de termo/opções
Taxa A.N.A.
Emolumentos
Total
0,00
0,00
0,14
0,14
D
D
D
D
Corretagem / Despesas
Clearing
Execução
Execução casa
ISS ( SÃO PAULO )
I.R.R.F. s/ operações, base 505,00 0,02
Outros
Total
0,00
0,00
0,00
0,00
0,00
0,00
D
D
D
D
D
D
Resumo dos Negócios
Especificações diversas
Vendas à vista
Compras à vista
Opções - vendas
Operações à termo
Valor das oper. c/ títulos públ. (v. nom.)
Valor das operações
0,00
505,00
3.000,00
0,00
0,00
0,00
0,00
3.505,00
A coluna Q indica liquidação no Agente do Qualificado.
(*) - Observações:
 2 - Corretora ou pessoa vinculada atuou na contraparte.
 # - Negócio direto
 8 - Liquidação Institucional.
 D - Day-Trade
 F - Cobertura
 B - Debêntures
A - Posição futuro T - Liquidação pelo Bruto
C - Clubes e fundos de Ações
P - Carteira Própria
H - Home Broker
X - Box
Y - Desmanche de Box
L - Precatório
I - POP
* O valor dos impostos pode variar conforme o município
Líquido para 2.496,01 D
Observação: (1) As operações a termo não são computadas no líquido da fatura
CLEAR CORRETORA - GRUPO XP
Debêntures
Opções - compras
 05/04/2018
Depositária
Taxa de Transferência de Ativos 0,00 D
@ - Taxa de Transferência de Ativo
Emitido eletronicamente: 10/05/2021 10:11:12
"""  # noqa: E501


def _unica(texto: str = NOTA) -> NotaExtraida:
    res = extrair_textos("nota.pdf", [texto])
    assert res.rejeitadas == []
    [nota] = res.notas
    return nota


def _resumo(texto: str = NOTA) -> ResumoBovespa:
    r = _unica(texto).resumo_bovespa
    assert r is not None
    return r


def test_layout_e_identificado_em_vez_do_perfil() -> None:
    assert classificar(NOTA).gramatica == "b3rv"


def test_cabecalho_conta_e_negocios() -> None:
    nota = _unica()
    assert (nota.corretora, nota.numero, nota.data_pregao) == ("CLEAR", "273207", date(2018, 4, 2))
    assert nota.conta == "765432"
    compra, venda = nota.negocios_bovespa
    assert (compra.cv, compra.mercado, compra.prazo) == ("C", "VISTA", "")
    assert compra.especificacao == "PETROBRAS PN N2"
    assert (compra.quantidade, compra.preco, compra.valor, compra.dc) == (
        Decimal(100), Decimal(30), Decimal(3000), "D"
    )  # fmt: skip
    assert (venda.cv, venda.mercado, venda.prazo) == ("V", "FRACIONARIO", "")
    assert (venda.especificacao, venda.observacao) == ("VALE ON NM", "#")


def test_resumo_em_blocos_fecha_sem_ajuste() -> None:
    r = _resumo()
    assert r.compras == Decimal(3000) and r.vendas == Decimal(505)
    assert r.valor_operacoes == Decimal(3505)
    assert r.valor_liquido_operacoes.assinado == Decimal(-2495)
    assert r.total_cblc.assinado == Decimal("-2495.87")
    assert r.total_bovespa.valor == Decimal("0.14")
    assert r.total_despesas.valor == Decimal(0)
    assert (r.irrf, r.irrf_base, r.irrf_no_liquido) == (Decimal("0.02"), Decimal(505), False)
    assert r.liquido.assinado == Decimal("-2496.01")
    assert r.data_liquidacao == date(2018, 4, 5)
    valores = {c.categoria: c.valor for c in r.custos if c.valor}
    assert valores == {
        CategoriaCusto.LIQUIDACAO: Decimal("0.87"), CategoriaCusto.EMOLUMENTOS: Decimal("0.14"),
    }  # fmt: skip
    assert not any(c.rotulo == ROTULO_AJUSTE for c in r.custos)


def test_cada_custo_vai_para_o_seu_rotulo() -> None:
    """Os valores vêm separados dos rótulos: a ordem é o que liga um ao outro."""
    trocado = NOTA.replace(
        "Total\n0,00\n0,00\n0,00\n0,00\n0,00\n0,00\n",
        "Total\n1,20\n0,30\n0,00\n0,15\n0,05\n1,70\n",
    ).replace("Líquido para 2.496,01 D", "Líquido para 2.497,71 D")
    custos = {c.rotulo: c.valor for c in _resumo(trocado).custos}
    assert custos["Clearing"] == Decimal("1.20")
    assert custos["Execução"] == Decimal("0.30")
    assert custos["Execução casa"] == Decimal(0)
    assert custos["ISS"] == Decimal("0.15")
    assert custos["Outros"] == Decimal("0.05")


@pytest.mark.parametrize(
    ("de", "para", "motivo"),
    [
        ("0,14\n0,14\nD", "0,14\nD", "valores soltos"),
        ("Emolumentos\n", "Taxa Nova\n", "custo desconhecido"),
        ("Líquido para 2.496,01 D\n", "", "Líquido para"),
        ("3.505,00\nA coluna", "A coluna", "resumo dos negócios"),
    ],
)
def test_nota_incompleta_nao_entra(de: str, para: str, motivo: str) -> None:
    assert de in NOTA
    res = extrair_textos("nota.pdf", [NOTA.replace(de, para)])
    assert res.notas == []
    [rejeitada] = res.rejeitadas
    assert motivo in rejeitada.motivo


def test_compras_e_vendas_trocadas_nao_passam_na_conciliacao() -> None:
    trocada = NOTA.replace("0,00\n505,00\n3.000,00\n", "0,00\n3.000,00\n505,00\n")
    res = extrair_textos("nota.pdf", [trocada])
    assert res.notas == []
    assert "compras" in res.rejeitadas[0].motivo


def test_cliente_impresso() -> None:
    c = cliente(NOTA)
    assert c is not None
    assert (c.cpf, c.nome) == ("00000000000", "FULANO DE TAL SILVA")


def test_corretora_sem_perfil_no_mesmo_layout_e_lida_com_aviso() -> None:
    outra = (
        NOTA.replace("CLEAR CORRETORA - GRUPO XP", "CORRETORA NOVA SA")
        .replace("clear", "nova")
        .replace("02.332.886/0011-78", "11.222.333/0001-81")
        .replace("308 - 9 765432", "999 - 9 765432")
    )
    res = extrair_textos("nota.pdf", [outra])
    assert res.rejeitadas == []
    [nota] = res.notas
    assert nota.corretora == "CORRETORA 999"
    assert any("layout padrão" in a for a in res.avisos)


def test_conferencia_do_gabarito_le_o_negocio() -> None:
    assert [(x.cv, x.valor) for x in linhas_bovespa(NOTA)] == [
        ("C", Decimal(3000)), ("V", Decimal(505)),
    ]  # fmt: skip


def test_copia_anonimizada_sem_dado_pessoal_e_fiel() -> None:
    feita = conferir("nota.pdf", [NOTA], MapaAnonimizacao.novo())
    assert feita.recusas == []
    [copia] = feita.textos
    for dado in (
        "FULANO", "SILVA", "000.000.000-00", "765432", "IMAGINARIA", "INVENTADO", "CIDADEFALSA",
        "13000-000", "98765", "4567890", "99888-7766",
    ):  # fmt: skip
        assert dado not in copia, dado
    assert _unica(copia).numero == "273207"
    assert _resumo(copia).liquido.assinado == Decimal("-2496.01")
