from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from coir import referencias
from coir.domain.nota import NotaExtraida, Segmento
from coir.erros import ConciliacaoError, ExtracaoError, ReferenciaError
from coir.parsing.conciliar import conciliar
from coir.parsing.extrair import ResultadoExtracao, completar_contas, extrair_textos
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
    pagina = classificar("BTG Pactual CTVM S.A.\nMercadoria: FCA - Com Garantia")
    assert pagina.tipo is TipoPagina.OPCAO_FLEXIVEL
    res = extrair_textos("flex.pdf", ["BTG Pactual CTVM S.A.\nMercadoria: FCA - Com Garantia"])
    assert res.notas == [] and res.rejeitadas == [] and "opção flexível" in res.avisos[0]


def test_conta_ausente_completada_so_se_unica_no_lote() -> None:
    nota = _nota("xp_2019_02_bmf_daytrade")
    sem = nota.model_copy(update={"conta": "", "numero": "2"})
    res = completar_contas(ResultadoExtracao(notas=[nota.model_copy(update={"conta": "77"}), sem]))
    assert [n.conta for n in res.notas] == ["77", "77"]
    assert "conta não impressa" in res.notas[1].avisos[0]
    duas = [nota.model_copy(update={"conta": c}) for c in ("77", "88")]
    res = completar_contas(ResultadoExtracao(notas=[*duas, sem]))
    assert len(res.notas) == 2 and "2 contas" in res.rejeitadas[0].motivo


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
    assert classes_das_series([opcoes])["BBASN380"] == "ON"


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
