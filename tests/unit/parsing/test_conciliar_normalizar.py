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
@pytest.mark.parametrize(
    ("de", "para", "campo"),
    [
        ("Corretagem 20,00 D", "Corretagem 21,00 D", "Total de despesas"),
        ("Emolumentos 5,03 D", "Emolumentos 5,04 D", "Total Bovespa / Soma"),
        ("18.686,92 D", "18.686,93 D", "líquido"),
        ("Taxa de liquidação 1,49 D", "Taxa de liquidação 1,48 D", "Total CBLC"),
    ],
)
def test_alterar_um_valor_impresso_quebra_a_conciliacao(de: str, para: str, campo: str) -> None:
    res = extrair_textos("x.pdf", _troca("agora_2008_04_bovespa", de, para))
    assert res.notas == []
    [rejeitada] = res.rejeitadas
    assert campo in rejeitada.motivo


def test_conciliar_bmf_detecta_custo_alterado() -> None:
    nota = _nota("xp_2019_02_bmf_daytrade")
    r = nota.resumo_bmf
    assert r is not None
    errada = nota.model_copy(
        update={"resumo_bmf": r.model_copy(update={"total_custos": Decimal("2.65")})}
    )
    with pytest.raises(ConciliacaoError, match="total de custos"):
        conciliar(errada)


def test_negocio_com_valor_diferente_de_quantidade_x_preco() -> None:
    res = extrair_textos("x.pdf", _troca("agora_2008_04_bovespa", "300 62,20", "300 62,21"))
    assert "quantidade x preço" in res.rejeitadas[0].motivo


# ------------------------------------------------------------------------------- páginas
def test_pagina_sem_perfil_e_opcao_flexivel() -> None:
    with pytest.raises(ExtracaoError, match="sem perfil"):
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


def test_marca_day_trade_sem_compra_e_venda() -> None:
    assert "day trade de GERDAU" in _rejeitada("agora_2008_04_bovespa", "N1 H 300", "N1 HD 300")


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


def test_ativo_sem_cadastro_falha_com_sugestao() -> None:
    nota = _nota("agora_2008_04_bovespa")
    neg = nota.negocios_bovespa[0].model_copy(update={"especificacao": "GERDAX PN N1"})
    estranha = nota.model_copy(update={"negocios_bovespa": (neg,)})
    with pytest.raises(ReferenciaError, match=r"parecidos: .*GGBR4"):
        operacoes(estranha, REFS, {})


def test_segmento_coerente() -> None:
    nota = _nota("agora_2008_04_bovespa")
    with pytest.raises(ValueError, match="incoerentes"):
        NotaExtraida.model_validate({**nota.model_dump(), "segmento": Segmento.BMF})
