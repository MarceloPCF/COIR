"""Futuros de commodities (café ICF e soja SJC, cotados em US$; milho e boi em R$): o resultado
vem dos ajustes em R$ impressos na nota, o multiplicador só entra no volume do rateio e na
conferência do câmbio implícito. Casos fechados à mão pelo especialista (19/10/2026).

Fontes: IN RFB 1.585/2015 art. 61 (resultado = soma dos ajustes no encerramento), art. 65 (15% e
20% day trade, inclusive bolsas de mercadorias e futuros); fichas B3 do Café Arábica 4/5 (100
sacas, US$/saca, entrega física) e do Minicontrato de Soja CME (450 sacas, US$/saca).
"""

from datetime import date
from decimal import Decimal
from typing import Literal

import pytest

from coir import referencias
from coir.domain.fiscal import Categoria, NotaMotor
from coir.domain.nota import Montante, NegocioBMF, NotaExtraida, ResumoBMF, Segmento
from coir.engine.futuros import movimentos_futuros
from coir.erros import ConciliacaoError, ExtracaoError, ReferenciaError
from coir.parsing.futuros import cambios, conversao_do_volume
from coir.servicos.apuracao import _ajustes  # pyright: ignore[reportPrivateUsage]

D = Decimal
REFS = referencias.carregar()
Z = Montante(valor=D(0))


def neg(
    cv: Literal["C", "V"],
    mercadoria: str,
    qtd: int,
    preco: str,
    tipo: str,
    valor: str,
    vencimento: date | None = None,
) -> NegocioBMF:
    v = D(valor)
    return NegocioBMF(
        cv=cv,
        mercadoria=mercadoria,
        vencimento=vencimento,  # type: ignore[arg-type]
        quantidade=D(qtd),
        preco=D(preco),
        tipo=tipo,
        valor=abs(v),
        dc="D" if v < 0 else "C",
        taxa_operacional=D(0),
    )


def nota(dia: date, *negocios: NegocioBMF, numero: str = "1") -> NotaExtraida:
    resumo = ResumoBMF(
        valor_negocios=Z,
        ajuste_posicao=Z,
        ajuste_day_trade=Z,
        custos=(),
        total_custos=D(0),
        impostos=D(0),
        irrf=D(0),
        irrf_day_trade=D(0),
        liquido=Z,
    )
    return NotaExtraida(
        corretora="XP",
        conta="1",
        numero=numero,
        data_pregao=dia,
        segmento=Segmento.BMF,
        arquivo="sintetica.pdf",
        paginas=(1,),
        negocios_bmf=negocios,
        resumo_bmf=resumo,
    )


VENC_ICF = date(2026, 5, 21)


def test_cafe_cambio_implicito_sai_da_linha_de_ajuste() -> None:
    """C1, 04/03/2026: AJUPOS 1 @ 299,00 (1.872,50 D) e venda 1 @ 301,00 (1.070,00 C):
    (301 - 299) x 1 x 100 = 200 US$ -> câmbio 1.070 / 200 = 5,35."""
    n = nota(
        date(2026, 3, 4),
        neg("C", "ICFK26", 1, "299.00", "AJUPOS", "-1872.50", VENC_ICF),
        neg("V", "ICFK26", 1, "301.00", "NORMAL", "1070.00", VENC_ICF),
    )
    assert cambios(n, REFS) == {"ICFK26": D("5.35")}
    assert conversao_do_volume(n, REFS) == {"ICFK26": D(1)}  # nota de uma moeda: sem conversão


def test_cafe_resultado_vem_dos_ajustes_em_reais() -> None:
    """C1 completo: abre em 03/03 (1.350,00 C, custos 3,00), ajuste em 04/03 (1.872,50 D) e
    fecha (1.070,00 C, custos 3,00): -525,50 + 1.067,00 = 541,50 (comum, 03/2026). O preço x
    multiplicador daria (301 - 300) x 100 = 100: o resultado não sai dele."""
    abre = nota(date(2026, 3, 3), neg("C", "ICFK26", 1, "300.00", "NORMAL", "1350.00", VENC_ICF))
    fecha = nota(
        date(2026, 3, 4),
        neg("C", "ICFK26", 1, "299.00", "AJUPOS", "-1872.50", VENC_ICF),
        neg("V", "ICFK26", 1, "301.00", "NORMAL", "1070.00", VENC_ICF),
        numero="2",
    )
    motor = [
        NotaMotor(
            id=i,
            titular=1,
            conta=1,
            corretora="XP",
            numero=n.numero,
            data=n.data_pregao,
            custos=D("3.00"),
            irrf=D(0),
            irrf_day_trade=D(0),
            ajustes=_ajustes(n, REFS),
        )
        for i, n in enumerate((abre, fecha), 1)
    ]
    [m] = movimentos_futuros(motor, date(2026, 3, 31))
    assert (m.categoria, m.resultado) == (Categoria.COMUM, D("541.50"))


def test_soja_dia_de_day_trade_sem_ajuste_confere_pelos_negocios_fechados() -> None:
    """Sem AJUPOS, os negócios que se fecham no dia dispensam o preço de ajuste: compra 2 @ 20,00 e
    venda 2 @ 20,30 (450 sacas) = 270 US$; impresso 1.458,00 -> câmbio 5,40."""
    n = nota(
        date(2026, 3, 5),
        neg("C", "SJCK26", 2, "20.00", "DAY TRADE", "486.00"),
        neg("V", "SJCK26", 2, "20.30", "DAY TRADE", "972.00"),
    )
    assert cambios(n, REFS) == {"SJCK26": D("5.4")}


def test_multiplicador_errado_e_conferido_pelo_cambio_implicito() -> None:
    """C7: no milho (R$), o câmbio tem de dar 1. A fixture real tem AJUPOS 71,76 e NORMAL 73,40
    com 738,00 D; com multiplicador 330 em vez de 450, a nota não entra."""
    n = nota(
        date(2024, 1, 10),
        neg("C", "CCMH24", 1, "71.76", "AJUPOS", "0"),
        neg("C", "CCMH24", 1, "73.40", "NORMAL", "-738.00"),
    )
    assert cambios(n, REFS) == {"CCMH24": D(1)}
    errado = REFS.model_copy(
        update={
            "contratos": tuple(
                c.model_copy(update={"multiplicador": D(330)}) if c.raiz == "CCM" else c
                for c in REFS.contratos
            )
        }
    )
    with pytest.raises(ConciliacaoError, match="multiplicador 330"):
        cambios(n, errado)


def test_nota_que_mistura_reais_e_dolares_converte_o_volume_do_rateio() -> None:
    """C5: WIN day trade (R$) + café (US$, câmbio 5,40 pela linha de ajuste). Custos de 10,00
    rateados pelo volume em R$: 130.000, 130.100 e 300 x 100 x 5,40 = 162.000 -> 3,08 + 3,08 + 3,84
    (sem a conversão o café ficaria com 1,03)."""
    n = nota(
        date(2026, 3, 10),
        neg("C", "WINJ26", 5, "130000", "DAY TRADE", "50.00"),
        neg("V", "WINJ26", 5, "130100", "DAY TRADE", "50.00"),
        neg("C", "ICFK26", 1, "302.50", "AJUPOS", "540.00", VENC_ICF),
        neg("C", "ICFK26", 1, "300.00", "NORMAL", "1350.00", VENC_ICF),
    )
    assert conversao_do_volume(n, REFS) == {"WINJ26": D(1), "ICFK26": D("5.4")}
    from coir.engine.rateio import ratear

    volumes = [a.volume for a in _ajustes(n, REFS) if a.volume]
    assert ratear(D("10.00"), volumes) == [D("3.08"), D("3.08"), D("3.84")]


def test_nota_mista_sem_cambio_calculavel_falha_alto() -> None:
    """Café aberto no dia (sem linha de ajuste) numa nota com WIN: sem câmbio para o rateio."""
    n = nota(
        date(2026, 3, 10),
        neg("C", "WINJ26", 5, "130000", "DAY TRADE", "50.00"),
        neg("V", "WINJ26", 5, "130100", "DAY TRADE", "50.00"),
        neg("C", "ICFK26", 1, "300.00", "NORMAL", "1350.00", VENC_ICF),
    )
    with pytest.raises(ConciliacaoError, match="câmbio implícito"):
        conversao_do_volume(n, REFS)


@pytest.mark.parametrize(
    ("negocio", "erro", "trecho"),
    [
        (neg("C", "ICFK26", 1, "300.00", "EXO", "10.00"), ExtracaoError, "tipo de negócio"),
        (neg("C", "ICFK26C300", 1, "3.00", "NORMAL", "10.00"), ExtracaoError, "opção sobre futuro"),
        (neg("C", "SOYK26", 1, "400.00", "NORMAL", "10.00"), ReferenciaError, "sem cadastro"),
    ],
)
def test_fora_do_escopo_falha_alto(negocio: NegocioBMF, erro: type[Exception], trecho: str) -> None:
    """C4 e C6: entrega física/exercício (tipo desconhecido), opção sobre futuro e contrato sem
    cadastro nunca entram calados."""
    with pytest.raises(erro, match=trecho):
        cambios(nota(date(2026, 3, 10), negocio), REFS)


def test_cafe_aberto_no_vencimento_avisa_entrega_fisica() -> None:
    """Café (liquidação física) com posição no vencimento: o resultado é o dos ajustes e o COIR
    avisa que a entrega está fora do escopo."""
    abre = nota(date(2026, 5, 4), neg("C", "ICFK26", 1, "300.00", "NORMAL", "1350.00", VENC_ICF))
    motor = NotaMotor(
        id=1,
        titular=1,
        conta=1,
        corretora="XP",
        numero="1",
        data=abre.data_pregao,
        custos=D(0),
        irrf=D(0),
        irrf_day_trade=D(0),
        ajustes=_ajustes(abre, REFS),
    )
    avisos: list[str] = []
    [m] = movimentos_futuros([motor], date(2026, 5, 31), avisos)
    assert m.resultado == D("1350.00") and m.data == VENC_ICF
    assert any("entrega física" in a for a in avisos)
