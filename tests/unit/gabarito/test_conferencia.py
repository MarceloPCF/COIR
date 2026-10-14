from datetime import date
from decimal import Decimal

from coir.gabarito.conferencia import (
    LinhaNegocio,
    ResultadoArquivo,
    conferir_textos,
    linhas_bmf,
    linhas_bovespa,
    relatorio,
    valores_impressos,
)
from coir.gabarito.modelo import Categoria, Gabarito, OperacaoGabarito


def test_linhas_bovespa_layouts() -> None:
    texto = "\n".join(
        [
            "BOVESPA C VISTA GERDAU PN N1 H 300 62,20 D 18.660,00",  # Ágora: D/C antes do valor
            "1-BOVESPA V VISTA COPASA ON NM @# 1.300 55,06 71.578,00 C",  # XP
            "1-BOVESPA V OPCAO DE VENDA 12/24 BBDCX512 PN 1000 1,25 1.250,00 C",  # BTG, sem milhar
            "1-BOVESPA V OPCAO DE COMPRA 06/24 VALEF730 ON D 400 1,35 540,00 C",  # obs. D
            "Resumo dos Negócios",
        ]
    )
    assert linhas_bovespa(texto) == [
        LinhaNegocio("C", Decimal(300), Decimal("62.20"), Decimal("18660.00")),
        LinhaNegocio("V", Decimal(1300), Decimal("55.06"), Decimal("71578.00")),
        LinhaNegocio("V", Decimal(1000), Decimal("1.25"), Decimal("1250.00")),
        LinhaNegocio("V", Decimal(400), Decimal("1.35"), Decimal("540.00")),
    ]


def test_linhas_bmf_ignora_ajuste_de_posicao() -> None:
    texto = "\n".join(
        [
            "C WING19 @13/02/2019 1 97.810,0000 DAYTRADE 25,00 C 0,20",
            "V WINM25 18/06/2025 5 137.945,0000 DAY TRADE 387,00 C 0,00",
            "C BGI G22 @25/02/2022 1 343,2500 NORMAL 115,50 D 14,00",
            "V CCM H24 @15/03/2024 6 64,30 NORMAL 0,00 24,00",  # abertura: 0,00 sem D/C
            "V CCMPPPP 0 0,0000 TX. PERMANÊNCIA 0,00 C 0,00",
            "V CCMU23 15/09/2023 2 58,8400 AJUPOS 1.593,00 D 0,00",
        ]
    )
    assert [(x.cv, x.quantidade, x.preco) for x in linhas_bmf(texto)] == [
        ("C", Decimal(1), Decimal("97810.0000")),
        ("V", Decimal(5), Decimal("137945.0000")),
        ("C", Decimal(1), Decimal("343.2500")),
        ("V", Decimal(6), Decimal("64.30")),
    ]


def test_valores_impressos() -> None:
    assert valores_impressos("0,00| 0,41 1,60 | D 1000,00 40,87 | C") == {
        Decimal("0.00"), Decimal("0.41"), Decimal("1.60"), Decimal("1000.00"), Decimal("40.87"),
    }  # fmt: skip


def _op(nota: str, cv: str, mercado: str, preco: str, total: str, custos: str) -> OperacaoGabarito:
    return OperacaoGabarito(
        categoria=Categoria.DAY_TRADE, linha=2, titular="TITULAR-1", corretora="XP",
        conta="XP-1", nota=nota, data=date(2026, 2, 2), cv=cv, papel="WING26",  # type: ignore[arg-type]
        mercado=mercado, preco=Decimal(preco), quantidade=Decimal(1), total=Decimal(total),
        custos=Decimal(custos), pm=None, irrf=Decimal(0), prazo=None, ativo_objeto=None,
    )  # fmt: skip


def _gabarito(ops: list[OperacaoGabarito]) -> Gabarito:
    return Gabarito(operacoes=ops, eventos=[], eventos_cadastro=[], posicoes=[], apuracao=[],
                    darfs=[], isentos=[], custodia_inicial=[], avisos=[])  # fmt: skip


PAGINA_BMF = "\n".join(
    [
        "C WIN G26 18/02/2026 1 183.685,00 DAY TRADE 30,20 C 0,20",
        "V WIN G26 18/02/2026 1 183.665,00 DAY TRADE 34,20 D 0,20",
        "NOTA DE NEGOCIAÇÃO Nr. nota",
        "409.950",
        "Ajuste day trade Total de custos operacionais",
        "0,00 0,04 0,00 | 4,00 | D 0,90 | D",
        "Total líquido da nota 4,94 | D",
    ]
)
OPS_BMF = [
    _op("409950", "C", "INDICE", "183685", "36737", "0.47"),
    _op("409950", "V", "INDICE", "183665", "36733", "0.47"),
]


def test_nota_bmf_dentro_do_pdf_de_acoes_e_conferida() -> None:
    """PDF misto: a nota BM&F está numa página do PDF Bovespa do mês (não há _BMF.pdf)."""
    textos = {"XP_NC_2026_02.pdf": ["Nr. nota 123\nsem negócios", PAGINA_BMF]}
    [arq] = conferir_textos(_gabarito(OPS_BMF), textos)
    assert arq.paginas_sem_nota == []
    [nota] = arq.notas
    assert nota.nota == "409950" and nota.ok
    assert {v.nome for v in nota.verificacoes} >= {"quantidade C", "ajuste day trade", "líquido"}


def test_pdf_do_segmento_tem_preferencia_e_pagina_usada_nao_e_reaproveitada() -> None:
    textos = {"XP_NC_2026_02.pdf": [PAGINA_BMF], "XP_NC_2026_02_BMF.pdf": [PAGINA_BMF]}
    resultados = {r.arquivo: r for r in conferir_textos(_gabarito(OPS_BMF), textos)}
    assert [n.nota for n in resultados["XP_NC_2026_02_BMF.pdf"].notas] == ["409950"]
    assert resultados["XP_NC_2026_02.pdf"].notas == []
    assert resultados["XP_NC_2026_02.pdf"].paginas_sem_nota == [1]


def test_nota_sem_nenhum_pdf_vira_divergencia() -> None:
    resultados = conferir_textos(_gabarito(OPS_BMF), {"XP_NC_2026_03.pdf": [PAGINA_BMF]})
    ausente = resultados[-1]
    assert ausente.arquivo == "XP_NC_2026_02_BMF.pdf (ausente)"
    [nota] = ausente.notas
    assert not nota.ok and nota.verificacoes[0].observacao == "PDF ausente"


def test_mesmo_numero_de_nota_em_datas_diferentes_fica_separado() -> None:
    outra = [o.model_copy(update={"data": date(2026, 2, 9)}) for o in OPS_BMF]
    paginas = [f"Data pregão 02/02/2026\n{PAGINA_BMF}", f"Data pregão 09/02/2026\n{PAGINA_BMF}"]
    [arq] = conferir_textos(_gabarito(OPS_BMF + outra), {"XP_NC_2026_02_BMF.pdf": paginas})
    assert [(n.data, n.ok) for n in arq.notas] == [("2026-02-02", True), ("2026-02-09", True)]


def test_mesmo_numero_e_data_em_contas_diferentes_nao_e_somado() -> None:
    outra_conta = [o.model_copy(update={"conta": "XP-2"}) for o in OPS_BMF]
    paginas = [PAGINA_BMF, PAGINA_BMF]
    [arq] = conferir_textos(_gabarito(OPS_BMF + outra_conta), {"XP_NC_2026_02_BMF.pdf": paginas})
    assert sorted(n.conta for n in arq.notas) == ["XP-1", "XP-2"]
    # o texto não diz qual página é de qual conta: não soma as duas, marca como não verificável
    assert [v.observacao for n in arq.notas for v in n.verificacoes] == [
        "mesmo número e data em outra conta"
    ] * 2
    assert all(v.ok is None for n in arq.notas for v in n.verificacoes)


def test_bmf_com_quantidade_zero_e_valor_zero_nao_quebram_nem_passam() -> None:
    """Quantidade 0 não pode derrubar a checagem de day trade; ajuste 0,00 não prova nada."""
    zerada = _op("409950", "C", "INDICE", "183685", "36737", "0.47").model_copy(
        update={"papel": "WINJ26", "quantidade": Decimal(0), "total": Decimal(0)}
    )
    empate = [o.model_copy(update={"total": Decimal(36737)}) for o in OPS_BMF]
    texto = {"XP_NC_2026_02_BMF.pdf": [PAGINA_BMF]}
    [arq] = conferir_textos(_gabarito([*empate, zerada]), texto)
    ajuste = next(v for v in arq.notas[0].verificacoes if v.nome == "ajuste day trade")
    assert (ajuste.ok, ajuste.observacao) == (None, "zero: não verificável")


def test_relatorio_conta_pdfs_lidos_ausentes_e_nao_lidos() -> None:
    resultados = conferir_textos(_gabarito(OPS_BMF), {"XP_NC_2026_03.pdf": [PAGINA_BMF]})
    resultados.append(ResultadoArquivo("BTG_NC_2023_11_Opcao-Flexivel.pdf", erro="fora do padrão"))
    md, csv = relatorio(resultados)
    assert "- PDFs: 1 lidos, 1 não lidos" in md
    assert "- BTG_NC_2023_11_Opcao-Flexivel.pdf: fora do padrão" in md
    assert "| XP_NC_2026_02_BMF.pdf (ausente) | 409950 |" in md
    assert csv[1][-1] == "PDF ausente"
