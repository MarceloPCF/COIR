"""Centavos de arredondamento não recusam a nota: viram a linha de custo "Ajuste de arredondamento"
e a nota fecha (decisão do usuário de 07/11/2026). Só o que passar do que o arredondamento explica
consome a tolerância de R$ 0,05 da nota, com aviso.

Números conferidos à mão na fixture agora_2008_04_bovespa: Emolumentos 5,03; Total Bovespa/Soma é a
soma do grupo; líquido 18.686,92 D (débito). Cada parcela impressa pode ter até R$ 0,005 de erro.
"""

from decimal import Decimal
from pathlib import Path

import pytest

from coir.domain.nota import NotaExtraida
from coir.erros import ConciliacaoError, ExtracaoError
from coir.parsing import bovespa
from coir.parsing.conciliar import ROTULO_AJUSTE, conciliar
from coir.parsing.extrair import extrair_textos

D = Decimal
PASTA = Path(__file__).parents[2] / "fixtures" / "notas"
AGORA = "agora_2008_04_bovespa"


def _textos(nome: str, de: str = "", para: str = "") -> list[str]:
    texto = (PASTA / f"{nome}.txt").read_text(encoding="utf-8")
    if de:
        assert de in texto
        texto = texto.replace(de, para, 1)
    return texto.split("\f")


def _ajustes(nota: NotaExtraida) -> list[tuple[str, Decimal]]:
    r = nota.resumo_bovespa or nota.resumo_bmf
    assert r is not None
    return [(c.grupo, c.valor) for c in r.custos if c.rotulo == ROTULO_AJUSTE]


def test_nota_que_fecha_nao_ganha_ajuste() -> None:
    [nota] = extrair_textos("x.pdf", _textos(AGORA)).notas
    assert _ajustes(nota) == []


@pytest.mark.parametrize(
    ("de", "para", "grupo", "valor"),
    [
        # impresso 5,04 com total 5,03+...: o total impresso é menor que a soma em 0,01
        ("Emolumentos 5,03 D", "Emolumentos 5,04 D", "bolsa", D("-0.01")),
        # líquido a pagar 1 centavo maior: mais débito para o cliente
        ("18.686,92 D", "18.686,93 D", "despesas", D("-0.01")),
        # a soma do CBLC calculada fica 1 centavo maior que a impressa
        ("Taxa de liquidação 1,49 D", "Taxa de liquidação 1,48 D", "cblc", D("0.01")),
    ],
)
def test_um_centavo_vira_ajuste_e_a_nota_entra_sem_aviso(
    de: str, para: str, grupo: str, valor: Decimal
) -> None:
    res = extrair_textos("x.pdf", _textos(AGORA, de, para))
    assert res.rejeitadas == []
    [nota] = res.notas
    [(g, v)] = _ajustes(nota)
    assert (g, abs(v)) == (grupo, abs(valor))
    assert res.avisos == []  # arredondamento não é erro: não consome a tolerância nem avisa


def test_o_ajuste_faz_os_custos_baterem_com_o_liquido_impresso() -> None:
    """Sem IRRF: líquido = valor líquido das operações - custos; com o ajuste, fecha exatamente."""
    [nota] = extrair_textos("x.pdf", _textos(AGORA, "18.686,92 D", "18.686,93 D")).notas
    r = nota.resumo_bovespa
    assert r is not None
    assert r.valor_liquido_operacoes.assinado - nota.custos_total == r.liquido.assinado


def test_conciliar_de_novo_da_o_mesmo_ajuste() -> None:
    [nota] = extrair_textos("x.pdf", _textos(AGORA, "18.686,92 D", "18.686,93 D")).notas
    assert conciliar(nota).resumo_bovespa == nota.resumo_bovespa


def test_alem_do_arredondamento_ate_cinco_centavos_entra_com_aviso() -> None:
    res = extrair_textos("x.pdf", _textos(AGORA, "18.686,92 D", "18.686,96 D"))
    assert len(res.notas) == 1
    assert any("tolerância de R$ 0,05" in a for a in res.avisos)


@pytest.mark.parametrize(
    ("de", "para", "campo"),
    [
        ("Corretagem 20,00 D", "Corretagem 21,00 D", "Total de despesas"),
        ("Emolumentos 5,03 D", "Emolumentos 5,23 D", "Total Bovespa / Soma"),
        ("18.686,92 D", "18.687,92 D", "líquido"),
        ("Taxa de liquidação 1,49 D", "Taxa de liquidação 1,29 D", "Total CBLC"),
    ],
)
def test_diferenca_maior_continua_recusando(de: str, para: str, campo: str) -> None:
    res = extrair_textos("x.pdf", _textos(AGORA, de, para))
    assert res.notas == []
    assert campo in res.rejeitadas[0].motivo


def test_ajuste_nao_deixa_o_custo_negativo() -> None:
    """Custo negativo não tem amparo (IN RFB 1.585/2015 art. 56 §3º) e o rateio o recusa."""
    [nota] = extrair_textos("x.pdf", _textos(AGORA)).notas
    r = nota.resumo_bovespa
    assert r is not None
    sem_custo = r.model_copy(
        update={
            "custos": (),
            "total_cblc": r.valor_liquido_operacoes,
            "total_bovespa": r.total_bovespa.model_copy(update={"valor": D(0)}),
            "total_despesas": r.total_despesas.model_copy(update={"valor": D(0)}),
            "liquido": r.valor_liquido_operacoes.model_copy(
                update={"valor": r.valor_liquido_operacoes.valor - D("0.01")}
            ),
        }
    )
    resultado = conciliar(nota.model_copy(update={"resumo_bovespa": sem_custo}))
    assert resultado.custos_total == 0
    assert any("custo zero" in a for a in resultado.avisos)


def test_irrf_de_centavos_nao_e_confundido_com_arredondamento_das_despesas() -> None:
    """IRRF de R$ 0,02 e total de despesas 1 centavo maior: sem bater exato, é arredondamento, e o
    IRRF não vai para as despesas nem para o líquido. O centavo do total se desfaz no líquido."""
    [nota] = extrair_textos("x.pdf", _textos(AGORA)).notas
    r = nota.resumo_bovespa
    assert r is not None
    total = r.total_despesas.model_copy(update={"valor": r.total_despesas.valor + D("0.01")})
    com_irrf = r.model_copy(update={"irrf": D("0.02"), "total_despesas": total})
    resultado = conciliar(nota.model_copy(update={"resumo_bovespa": com_irrf})).resumo_bovespa
    assert resultado is not None and not resultado.irrf_no_liquido
    assert sum(c.valor for c in resultado.custos) == sum(c.valor for c in r.custos)


def test_bmf_custo_de_centavo_vira_ajuste_e_grande_recusa() -> None:
    [nota] = extrair_textos("x.pdf", _textos("xp_2019_02_bmf_daytrade")).notas
    r = nota.resumo_bmf
    assert r is not None
    perto = nota.model_copy(
        update={"resumo_bmf": r.model_copy(update={"total_custos": r.total_custos + D("0.01")})}
    )
    # o líquido impresso não mudou: a diferença do total se desfaz no líquido
    assert sum(v for _, v in _ajustes(conciliar(perto))) == 0
    longe = nota.model_copy(
        update={"resumo_bmf": r.model_copy(update={"total_custos": r.total_custos + D(1)})}
    )
    with pytest.raises(ConciliacaoError, match=r"total de custos|líquido"):
        conciliar(longe)


# --------------------------------------------------------------------------- linha do negócio
@pytest.mark.parametrize(
    ("q", "p", "v"),
    [
        ("3", "17,45", "52,34"),  # preço médio impresso: 3 x 17,45 = 52,35
        ("100", "13,73", "1.373,40"),  # 13,734 de preço médio, impresso 13,73
        ("10", "24,54", "2,45"),  # RAIZ55: fator 100
        ("800", "24,53", "196,24"),  # RAIZ55: fator 100
        ("1.000", "12,34", "12,34"),  # lote de mil sem o asterisco
        ("200", "4,11", "822,00"),
    ],
)
def test_linha_aceita_arredondamento_do_preco_e_fator_de_cotacao(q: str, p: str, v: str) -> None:
    [n] = bovespa.negocios(f"1-BOVESPA C VISTA VALE ON NM {q} {p} {v} D", "xp")
    assert n.valor == D(v.replace(".", "").replace(",", "."))


@pytest.mark.parametrize(
    ("q", "p", "v"),
    [("100", "10,00", "1.100,00"), ("300", "62,21", "18.660,00"), ("3", "17,45", "52,30"),
     ("1.000", "10,00", "100,50")],  # perto do fator 100, sem bater exato: não é lote
)  # fmt: skip
def test_linha_que_nem_o_arredondamento_explica_recusa(q: str, p: str, v: str) -> None:
    with pytest.raises(ExtracaoError, match="quantidade x preço"):
        bovespa.negocios(f"1-BOVESPA C VISTA VALE ON NM {q} {p} {v} D", "xp")
