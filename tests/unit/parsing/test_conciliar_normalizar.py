from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from coir import referencias
from coir.domain.nota import NotaExtraida, Segmento
from coir.erros import ConciliacaoError, ExtracaoError, ReferenciaError
from coir.parsing.conciliar import conciliar
from coir.parsing.extrair import ResultadoExtracao, extrair_textos, sem_duplicadas
from coir.parsing.normalizar import classes_das_series, operacoes
from coir.parsing.paginas import TipoPagina, classificar

PASTA = Path(__file__).parents[2] / "fixtures" / "notas"


def _nota(nome: str) -> NotaExtraida:
    paginas = (PASTA / f"{nome}.txt").read_text(encoding="utf-8").split("\f")
    [nota] = extrair_textos(f"{nome}.pdf", paginas).notas
    return nota


def _troca(texto_nome: str, de: str, para: str) -> list[str]:
    texto = (PASTA / f"{texto_nome}.txt").read_text(encoding="utf-8")
    assert de in texto
    return texto.replace(de, para, 1).split("\f")


# ------------------------------------------------------------------------------- conciliação
def test_conciliar_bmf_detecta_custo_alterado() -> None:
    nota = _nota("xp_2019_02_bmf_daytrade")
    r = nota.resumo_bmf
    assert r is not None
    errada = nota.model_copy(
        update={"resumo_bmf": r.model_copy(update={"total_custos": r.total_custos + Decimal(1)})}
    )
    with pytest.raises(ConciliacaoError, match=r"total de custos|líquido"):
        conciliar(errada)


def test_negocio_com_valor_diferente_de_quantidade_x_preco() -> None:
    res = extrair_textos("x.pdf", _troca("agora_2008_04_bovespa", "300 62,20", "300 62,21"))
    assert "quantidade x preço" in res.rejeitadas[0].motivo


# ------------------------------------------------------------------------------- páginas
def test_pagina_sem_perfil_e_opcao_flexivel() -> None:
    with pytest.raises(ExtracaoError, match="corretora não identificada"):
        classificar("NOTA DE CORRETAGEM\nCorretora Desconhecida S.A.")
    flexivel = "BTG Pactual CTVM S.A.\nFCA - Com Garantia C 1 LIQ.\nMercadoria: Nat. Oper.:"
    assert classificar(flexivel).tipo is TipoPagina.OPCAO_FLEXIVEL
    res = extrair_textos("flex.pdf", [flexivel])
    assert res.notas == [] and res.rejeitadas == [] and "opção flexível" in res.avisos[0]


def _rejeitada(texto_nome: str, de: str, para: str) -> str:
    res = extrair_textos("x.pdf", _troca(texto_nome, de, para))
    assert res.notas == []
    [rejeitada] = res.rejeitadas
    return rejeitada.motivo


def test_conta_bmf_nao_impressa_rejeita_sem_palpite() -> None:
    texto = (PASTA / "xp_2019_02_bmf_daytrade.txt").read_text(encoding="utf-8").rstrip()
    sem_conta = texto.rsplit("\n", 1)[0]  # a última linha é a conta
    res = extrair_textos("x.pdf", [sem_conta])
    assert res.notas == [] and "não impressa" in res.rejeitadas[0].motivo


@pytest.mark.parametrize(
    "rodape",
    [
        "Atendimento: 0800 000 0000 Ouvidoria: 0800 000 1443",  # telefone no fim (até 2024)
        "Grupo Financeiro Banco Exemplo S/A, Rua Exemplo 0, Cidade, UF",  # sem número (2025+)
    ],
)
def test_conta_bmf_btg_vem_depois_de_codigo_cliente(rodape: str) -> None:
    """Na BTG a conta fica no fim da linha seguinte a "... Código cliente" (depois do endereço);
    a última linha da página é o rodapé, que não é a conta."""
    texto = (PASTA / "btg_2023_11_bmf.txt").read_text(encoding="utf-8").rstrip()
    texto = texto.replace(
        "Código cliente\n", "RUA EXEMPLO, 1 - APTO 2 Código cliente\nCIDADE - UF 1234567\n", 1
    )
    res = extrair_textos("x.pdf", [f"{texto}\n{rodape}"])
    assert res.rejeitadas == [] and [n.conta for n in res.notas] == ["1234567"]


def test_cabecalho_corrompido_vira_rejeitada() -> None:
    texto = (PASTA / "xp_2026_08_bovespa.txt").read_text(encoding="utf-8")
    res = extrair_textos("x.pdf", [texto.replace("Folha\n1", "Folha\nX")])
    assert res.notas == [] and "folha inválido" in res.rejeitadas[0].motivo


def test_folhas_fora_de_sequencia_nao_formam_nota() -> None:
    texto = (PASTA / "xp_2016_11_bovespa_antigo_multipagina.txt").read_text(encoding="utf-8")
    folha1, folha2 = texto.split("\f")
    res = extrair_textos("x.pdf", [folha1, folha2.replace("Folha\n2", "Folha\n3")])
    assert res.notas == [] and any("sequência" in r.motivo for r in res.rejeitadas)


def test_linha_bmf_fora_do_padrao_falha_alto() -> None:
    motivo = _rejeitada("xp_2019_02_bmf_daytrade", "DAYTRADE 25,00 C 0,20", "DAYTRADE 25,00 C")
    assert "BM&F fora do padrão" in motivo


def test_marca_day_trade_sem_compra_e_venda_avisa_e_nao_recusa() -> None:
    """O outro lado pode estar em outra nota do dia; a apuração confere as quantidades."""
    res = extrair_textos("x.pdf", _troca("agora_2008_04_bovespa", "N1 H 300", "N1 HD 300"))
    assert res.rejeitadas == [] and len(res.notas) == 1
    assert any("day trade de GERDAU" in a for a in res.avisos)


def test_irrf_impresso_e_nao_lido_falha_alto() -> None:
    motivo = _rejeitada("rico_2015_11_bovespa", "baseR$10.660,00 0,53", "baseR$ 10.660,00 x0,53")
    assert "I.R.R.F." in motivo


def test_nome_terminado_em_letras_de_marca_nao_vira_observacao() -> None:
    from coir.parsing.bovespa import negocios

    linha = "1-BOVESPA C VISTA ALLIANCE DAT 100 10,00 1.000,00 D"
    [n] = negocios(linha, "xp")
    assert (n.especificacao, n.observacao, n.day_trade) == ("ALLIANCE DAT", "", False)
    [m] = negocios("1-BOVESPA V VISTA COPASA ON NM D#2 100 10,00 1.000,00 C", "xp")
    assert (m.especificacao, m.observacao, m.day_trade) == ("COPASA ON NM", "D#2", True)


def test_nota_repetida() -> None:
    nota = _nota("xp_2026_08_bovespa")
    outra = nota.model_copy(update={"arquivo": "y.pdf"})
    res = sem_duplicadas(ResultadoExtracao(notas=[nota, outra]))
    assert res.notas == [nota] and "repetida (idêntica)" in res.avisos[0]
    r = nota.resumo_bovespa
    assert r is not None
    diferente = outra.model_copy(update={"resumo_bovespa": r.model_copy(update={"irrf": 0})})
    res = sem_duplicadas(ResultadoExtracao(notas=[nota, diferente]))
    assert res.notas == [] and len(res.rejeitadas) == 2


def test_opcao_flexivel_exige_modalidade_no_inicio_da_linha() -> None:
    texto = "BTG Pactual CTVM S.A.\nMercadoria: Nat. Oper.:\nVALOR DE FCA ESTIMADO"
    assert classificar(texto).tipo is TipoPagina.NOTA


# ------------------------------------------------------------------------------- normalização
REFS = referencias.carregar()


def test_nome_de_pregao_vira_ticker_e_sufixos_saem() -> None:
    ops = operacoes(_nota("xp_2026_08_bovespa"), REFS, {})
    assert {o.papel for o in ops} == {"CSMG3", "OPCT3"}
    assert {o.mercado for o in ops} == {"VISTA"}
    ops = operacoes(_nota("agora_2008_04_bovespa"), REFS, {})
    assert [(o.papel, o.cv, o.total) for o in ops] == [("GGBR4", "C", Decimal("18660.00"))]


def test_opcoes_e_exercicio_pelo_ativo_objeto() -> None:
    opcoes = _nota("btg_2023_01_bovespa_opcoes")
    ops = operacoes(opcoes, REFS, {})
    assert [(o.papel, o.mercado, o.ativo_objeto) for o in ops][:2] == [
        ("BBASN380", "PUT", "BBAS3"),
        ("BBDCN128", "PUT", "BBDC4"),
    ]
    assert ops[-1].ativo_objeto == "TAEE11"
    exercicio = _nota("btg_2023_01_bovespa_exercicio")
    # sem classe na linha: classe da abertura da série; sem série conhecida, data/opcoes_raiz.csv
    [ex] = operacoes(exercicio, REFS, {"KLBNM199": "ON"})
    assert (ex.papel, ex.mercado, ex.ativo_objeto) == ("KLBN3", "EXERC PUT", "KLBNM199")
    assert classes_das_series([opcoes], REFS)["BBASN380"] == "ON"


def test_bmf_total_com_multiplicador_e_ajuste_de_posicao_fora() -> None:
    ops = operacoes(_nota("xp_2019_02_bmf_daytrade"), REFS, {})
    assert {(o.papel, o.mercado, o.categoria) for o in ops} == {("WING19", "INDICE", "DAY_TRADE")}
    assert ops[0].total == Decimal("97810.0000") * Decimal("0.2")
    ajupos = operacoes(_nota("xp_2024_01_bmf_novo_ajupos_negativo"), REFS, {})
    assert all(o.quantidade > 0 for o in ajupos) and {o.mercado for o in ajupos} == {"MILHO"}


def test_troca_de_ticker_por_data() -> None:
    assert REFS.ticker_na_data("JSLG3", date(2020, 9, 18)) == "SIMH3"
    assert REFS.ticker_na_data("JSLG3", date(2020, 9, 19)) == "JSLG3"


def test_ativo_sem_cadastro_diz_o_que_fazer() -> None:
    nota = _nota("agora_2008_04_bovespa")
    neg = nota.negocios_bovespa[0].model_copy(update={"especificacao": "NOME QUE NAO EXISTE ON"})
    estranha = nota.model_copy(update={"negocios_bovespa": (neg,)})
    with pytest.raises(ReferenciaError, match="Atualize as tabelas em Configurações") as erro:
        operacoes(estranha, REFS, {})
    assert "Parecidos" not in str(erro.value)  # sugestão fraca já levou ao ativo errado


@pytest.mark.parametrize(
    ("especificacao", "dia", "ticker"),
    [
        ("PETZCOBASI ON NM", date(2026, 10, 26), "AUAU3"),
        ("PETZCOBASI ON ATZ NM", date(2026, 10, 14), "AUAU3"),
        ("PETZ ON NM", date(2025, 6, 2), "AUAU3"),  # segue a troca de código
        ("RAIZEN PN N2", date(2026, 10, 30), "RAIZ55"),
        ("RAIZEN PNA P N2", date(2026, 11, 3), "RAIZ55"),
        ("XPETFCHINA CI", date(2025, 6, 2), "XINA11"),
        ("INVESTO NFTS CI", date(2025, 6, 2), "NFTS11"),
        ("INVESTO HODL CI", date(2026, 1, 5), "HODL11"),
    ],
)
def test_nomes_de_pregao_dos_casos_reais(especificacao: str, dia: date, ticker: str) -> None:
    """Como a B3 os publica no COTAHIST (fonte pública), pela data do negócio."""
    nota = _nota("agora_2008_04_bovespa")
    neg = nota.negocios_bovespa[0].model_copy(update={"especificacao": especificacao})
    nova = nota.model_copy(update={"negocios_bovespa": (neg,), "data_pregao": dia})
    assert operacoes(nova, REFS, {})[0].papel == ticker


def test_nome_usado_por_dois_tickers_vale_o_do_periodo() -> None:
    refs = REFS.model_copy(update={"nomes_historicos": {
        "nome": (("2010-01-01", "2012-12-31", "AAAA3"), ("2015-01-01", "2020-12-31", "BBBB3")),
        "juntos": (("2010-01-01", "2012-12-31", "AAAA3"), ("2011-01-01", "2020-12-31", "BBBB3")),
    }})  # fmt: skip
    assert refs.nome_na_data("nome", date(2011, 5, 5)) == "AAAA3"
    assert refs.nome_na_data("nome", date(2016, 5, 5)) == "BBBB3"
    assert refs.nome_na_data("nome", date(2013, 2, 1)) == "AAAA3"  # sem negócio: o mais próximo
    assert refs.nome_na_data("juntos", date(2011, 5, 5)) is None  # sem palpite
    assert refs.nome_na_data("outro", date(2011, 5, 5)) is None
    assert refs.nome_na_data("nome", date(2024, 5, 5)) is None  # longe de todos: sem palpite


def test_segmento_coerente() -> None:
    nota = _nota("agora_2008_04_bovespa")
    with pytest.raises(ValueError, match="incoerentes"):
        NotaExtraida.model_validate({**nota.model_dump(), "segmento": Segmento.BMF})


def test_termo_fora_do_liquido_mas_no_valor_das_operacoes() -> None:
    """Compra a termo: entra em "Operações à termo" e no valor das operações, não no líquido."""
    nota = _nota("xp_2026_08_bovespa")
    r = nota.resumo_bovespa
    assert r is not None
    termo = nota.negocios_bovespa[0].model_copy(
        update={"cv": "C", "mercado": "TERMO", "prazo": "90", "dc": "D"}
    )
    com_termo = nota.model_copy(
        update={
            "negocios_bovespa": (*nota.negocios_bovespa, termo),
            "resumo_bovespa": r.model_copy(
                update={"termo": termo.valor, "valor_operacoes": r.valor_operacoes + termo.valor}
            ),
        }
    )
    assert conciliar(com_termo).resumo_bovespa is not None
    errada = com_termo.model_copy(update={"resumo_bovespa": r.model_copy(update={"termo": 0})})
    with pytest.raises(ConciliacaoError, match="termo"):
        conciliar(errada)


def test_termo_guarda_o_prazo_do_contrato() -> None:
    nota = _nota("agora_2008_04_bovespa")
    neg = nota.negocios_bovespa[0].model_copy(update={"mercado": "TERMO", "prazo": "90"})
    [op] = operacoes(nota.model_copy(update={"negocios_bovespa": (neg,)}), REFS, {})
    assert (op.mercado, op.prazo_termo_dias) == ("TERMO", 90)
    [vista] = operacoes(nota, REFS, {})
    assert vista.prazo_termo_dias is None


@pytest.mark.parametrize("classe", ["PN", "PNA"])
def test_exercicio_de_opcao_depois_da_troca_da_classe(classe: str) -> None:
    """RAIZ PN negocia como PNA em RAIZ55 desde 03/11/2026 (B3 OC 022/2026-VPE)."""
    nota = _nota("agora_2008_04_bovespa")
    neg = nota.negocios_bovespa[0].model_copy(
        update={"especificacao": f"RAIZQ30E {classe}", "mercado": "EXERC OPC VENDA"}
    )
    nova = nota.model_copy(update={"negocios_bovespa": (neg,), "data_pregao": date(2026, 11, 20)})
    assert operacoes(nova, REFS, {})[0].papel == "RAIZ55"


DIREITOS = REFS.model_copy(update={"nomes_historicos": {
    "empresaxon": (("2010-01-04", "2026-11-09", "EMPX3"),),
    "empresaxdirord": (
        ("2020-07-14", "2020-08-07", "EMPX1"), ("2023-03-01", "2023-03-20", "EMPX1"),
    ),
    "empresaxdirpre": (("2023-03-01", "2023-03-20", "EMPX2"),),
    "fiifundoyci": (("2018-07-31", "2026-11-09", "FNDY11"),),
    "fiifundoydir": (("2019-01-21", "2019-02-08", "FNDY12"),),
}})  # fmt: skip


@pytest.mark.parametrize(
    ("especificacao", "dia", "ticker"),
    [
        ("EMPRESA X DO 6,93", date(2020, 7, 20), "EMPX1"),  # um direito só no período
        ("FII FUNDO Y DM 86,50", date(2019, 1, 25), "FNDY12"),  # a letra não precisa ser conhecida
        ("EMPRESA X DO 12,00", date(2023, 3, 10), "EMPX1"),  # duas classes no dia: a letra decide
        ("EMPRESA X DP 12,00", date(2023, 3, 10), "EMPX2"),
    ],
)
def test_direito_de_subscricao_com_o_preco_no_nome(
    especificacao: str, dia: date, ticker: str
) -> None:
    """A nota imprime o direito como "<nome> D<classe> <preço de subscrição>"; a B3 o lista como
    "<nome> DIR ORD" (ou PRE, ou só DIR). Vale o direito daquele nome negociado na data."""
    nota = _nota("agora_2008_04_bovespa")
    neg = nota.negocios_bovespa[0].model_copy(update={"especificacao": especificacao})
    nova = nota.model_copy(update={"negocios_bovespa": (neg,), "data_pregao": dia})
    assert operacoes(nova, DIREITOS, {})[0].papel == ticker


@pytest.mark.parametrize(
    ("especificacao", "dia"),
    [
        ("EMPRESA X DO 6,93", date(2021, 5, 5)),  # nenhum direito negociado na data
        ("EMPRESA X DM 12,00", date(2023, 3, 10)),  # duas classes e a letra não decide
        ("OUTRA EMPRESA DO 6,93", date(2020, 7, 20)),  # empresa sem direito na tabela
    ],
)
def test_direito_sem_um_unico_candidato_nao_e_chutado(especificacao: str, dia: date) -> None:
    nota = _nota("agora_2008_04_bovespa")
    neg = nota.negocios_bovespa[0].model_copy(update={"especificacao": especificacao})
    nova = nota.model_copy(update={"negocios_bovespa": (neg,), "data_pregao": dia})
    with pytest.raises(ReferenciaError, match="não encontrado"):
        operacoes(nova, DIREITOS, {})


def test_nome_de_dois_codigos_na_data_diz_quais_sao() -> None:
    """Dois BDR da mesma companhia com o mesmo nome de pregão no período: a nota é recusada dizendo
    os dois códigos, em vez de "ativo não encontrado"."""
    refs = REFS.model_copy(update={"nomes_historicos": {
        "empresaydr1": (
            ("2021-10-04", "2026-11-09", "EMPY31"), ("2023-07-03", "2024-09-12", "EMPY20"),
        ),
    }})  # fmt: skip
    nota = _nota("agora_2008_04_bovespa")
    neg = nota.negocios_bovespa[0].model_copy(update={"especificacao": "EMPRESA Y DR1"})
    antes = nota.model_copy(update={"negocios_bovespa": (neg,), "data_pregao": date(2022, 5, 5)})
    assert operacoes(antes, refs, {})[0].papel == "EMPY31"  # fora do período em comum: decide
    durante = nota.model_copy(update={"negocios_bovespa": (neg,), "data_pregao": date(2023, 8, 1)})
    with pytest.raises(ReferenciaError, match=r"mais de um código \(EMPY20, EMPY31\)"):
        operacoes(durante, refs, {})
