"""Histórico de eventos (FRE da CVM + COTAHIST da B3) e a junção com a consulta de empresas da B3:
fator só de fração simples, data com só quando a cotação confirma, divergência entre fontes nunca
resolvida por palpite, e o arquivo curado dos avisos da CVM."""

import io
import zipfile
from datetime import date, timedelta
from fractions import Fraction
from pathlib import Path

import pytest

from coir.referencias import evento_b3_valido
from tools.referencias import gerar, historico
from tools.referencias.gerar import EventoB3, Resumo
from tools.referencias.historico import LinhaFre, Pregao

CNPJ = "61532644000115"
CAB_FRE = ("CNPJ_Companhia;Data_Referencia;Versao;ID_Documento;Nome_Companhia;"
           "ID_Capital_Social_Desdobramento;Data_Aprovacao;Tipo_Evento;"
           "Quantidade_Acoes_Ordinarias_Antes_Aprovacao;"
           "Quantidade_Acoes_Preferenciais_Antes_Aprovacao;Quantidade_Total_Acoes_Antes_Aprovacao;"
           "Quantidade_Acoes_Ordinarias_Depois_Aprovacao;"
           "Quantidade_Acoes_Preferenciais_Depois_Aprovacao;"
           "Quantidade_Total_Acoes_Depois_Aprovacao")  # fmt: skip


def _zip(arquivos: dict[str, str]) -> bytes:
    saida = io.BytesIO()
    with zipfile.ZipFile(saida, "w") as z:
        for nome, texto in arquivos.items():
            z.writestr(nome, texto.encode("latin-1"))
    return saida.getvalue()


def test_fre_por_especie_com_as_classes_de_preferencial() -> None:
    base = (f"{CAB_FRE}\n61.532.644/0001-15;2016-01-01;3;65000;ITAUSA;9;2016-04-29;Bonificação;"
            "1000;2000;3000;1100;2200;3300\n"
            "61.532.644/0001-15;2016-01-01;3;65000;ITAUSA;10;2016-05-02;Ajuste;"
            "1;1;2;1;1;2\n")  # fmt: skip
    classes = ("CNPJ_Companhia;Data_Referencia;Versao;ID_Documento;Nome_Companhia;"
               "ID_Capital_Social_Desdobramento;Tipo_Classe_Acao_Preferencial;"
               "Quantidade_Acoes_Antes_Aprovacao;Quantidade_Acoes_Depois_Aprovacao\n"
               "61.532.644/0001-15;2016-01-01;3;65000;ITAUSA;9;Preferencial Classe B;"
               "500;550\n")  # fmt: skip
    linhas = historico.ler_fre(_zip({
        "fre_cia_aberta_capital_social_desdobramento_2016.csv": base,
        "fre_cia_aberta_capital_social_desdobramento_classe_acao_2016.csv": classes,
    }))  # fmt: skip
    assert [(x.classe, x.antes, x.depois) for x in linhas] == [("NOR", 1000, 1100),
                                                              ("NPR", 2000, 2200),
                                                              ("NPB", 500, 550)]  # fmt: skip
    assert {(x.cnpj, x.aprovado_em, x.tipo, x.documento) for x in linhas} == {
        (CNPJ, "2016-04-29", "BONIFICACAO", "65000")
    }  # tipo desconhecido fica de fora
    assert historico.ler_fre(_zip({"fre_cia_aberta_2024.csv": "x"})) == []  # FRE novo: sem tabela


@pytest.mark.parametrize(("antes", "depois", "fator"), [
    (3549302243, 3726767355, Fraction(21, 20)),  # ITSA 2023: 5%, fração vendida em leilão
    (27850069, 30381893, Fraction(12, 11)),  # UNIP 2018: 9,0909%
    (2467890331, 82263011, Fraction(1, 30)),  # IRBR 2022: grupamento 30:1
    (29953355, 89860065, Fraction(3)),  # KEPL 2022: desdobramento 1:3
    (886959131950, 88695913, Fraction(1, 10000)),  # TELB 2010
    (1651930000, 1718007200, Fraction(26, 25)),  # RADL 2023: 4%, percentual em meio ponto
    (5000000, 5491000, None),  # 9,82%: não é fração simples (ruído): revisão
    (4435106575, 4870579247, None),  # BBDC 2021: 9,8187%, perto de 10% mas não é
    (100, 100, None),
    (0, 10, None),
])  # fmt: skip
def test_fator_so_de_fracao_simples(antes: int, depois: int, fator: Fraction | None) -> None:
    assert historico.fator_simples(antes, depois) == fator


def test_fator_como_a_b3_escreve() -> None:
    assert historico.fator_b3("GRUPAMENTO", Fraction(1, 30)) == "0.03333333334"  # como a B3
    assert historico.fator_b3("BONIFICACAO", Fraction(4, 3)) == "33.333333334"  # piso não perde
    assert historico.fator_b3("BONIFICACAO", Fraction(21, 20)) == "5"
    assert historico.fator_b3("DESDOBRAMENTO", Fraction(3)) == "200"
    assert historico.fator_b3("GRUPAMENTO", Fraction(1, 40)) == "0.025"
    assert gerar.multiplicador("DESDOBRAMENTO", "200") == 3
    assert gerar.multiplicador("GRUPAMENTO", "0.1") == Fraction(1, 10)


def _linha_cotahist(dia: str, isin: str, distribuicao: str, preco: float, ticker: str = "X") -> str:
    linha = [" "] * 245
    campos = {0: "01", 2: dia, 10: "02", 12: ticker.ljust(12), 24: "010",
              108: f"{round(preco * 100):013d}", 210: "0000001", 230: isin,
              242: distribuicao}  # fmt: skip
    for inicio, texto in campos.items():
        linha[inicio : inicio + len(texto)] = texto
    return "".join(linha)


def _cotahist(linhas: list[str]) -> bytes:
    return _zip({"COTAHIST_A2016.TXT": "\r\n".join(["00COTAHIST.2016", *linhas]) + "\r\n"})


ON, PN, UNIT = "BRITSAACNOR0", "BRITSAACNPR7", "BRITSACDAM10"
DIAS = ["20160427", "20160428", "20160429", "20160502", "20160503", "20160504", "20160505",
        "20160506"]  # fmt: skip


def _series(precos: dict[str, list[tuple[str, float]]]) -> tuple[historico.Series, list[str]]:
    """Pregões dos dias de DIAS: (distribuição, preço) por ISIN; '' = sem negócio no dia."""
    linhas = [_linha_cotahist(d, isin, dist, p)
              for isin, valores in precos.items()
              for d, (dist, p) in zip(DIAS, valores, strict=True) if dist]  # fmt: skip
    series: historico.Series = {}
    pregoes: set[str] = set()
    historico.ler_cotahist(_cotahist(linhas), series, pregoes, {"ITSA"})
    return series, sorted(pregoes)


def test_cotahist_so_acoes_e_units_do_lote_padrao() -> None:
    series, pregoes = _series({ON: [("100", 10.0)] * 8, UNIT: [("100", 50.0)] * 8})
    assert set(series) == {ON, UNIT} and pregoes == DIAS
    assert series[ON][0] == Pregao("20160427", "100", 10.0)


def test_data_com_pela_unica_mudanca_coerente_com_o_fator() -> None:
    # dividendo (distribuição muda, preço quase igual) e depois a bonificação de 10% (cai ~9%)
    precos = [("100", 10.0), ("100", 10.0), ("101", 9.95), ("101", 10.0), ("102", 9.1),
              ("102", 9.1), ("102", 9.2), ("102", 9.2)]  # fmt: skip
    series, pregoes = _series({ON: precos})
    assert historico.data_com(series[ON], pregoes, "2016-04-28", Fraction(11, 10)) == (
        "2016-05-02",
        "",
    )  # data ex 03/05: data com é o pregão anterior
    # aprovação depois do evento: nenhuma mudança coerente na janela
    com, motivo = historico.data_com(series[ON], pregoes, "2016-05-03", Fraction(11, 10))
    assert com is None and motivo.startswith("nenhuma mudança")
    # fator pequeno demais (2%): a cotação do dia não distingue
    assert historico.data_com(series[ON], pregoes, "2016-04-28", Fraction(51, 50))[0] is None
    # antes do primeiro negócio: não atinge investidor da B3
    assert historico.data_com(series[ON], pregoes, "2016-04-01", Fraction(11, 10)) == (
        None,
        "antes da listagem",
    )


def test_data_com_ambigua_ou_sem_negocio_vai_para_revisao() -> None:
    duas = [("100", 10.0), ("100", 10.0), ("101", 9.1), ("101", 9.1), ("102", 8.3),
            ("102", 8.3), ("102", 8.3), ("102", 8.3)]  # fmt: skip
    series, pregoes = _series({ON: duas})
    com, motivo = historico.data_com(series[ON], pregoes, "2016-04-27", Fraction(11, 10))
    assert com is None and motivo.startswith("mais de uma data ex possível")
    sem_negocio = [("100", 10.0), ("100", 10.0), ("", 0), ("101", 9.1), ("101", 9.1),
                   ("101", 9.1), ("101", 9.1), ("101", 9.1)]  # fmt: skip
    series, pregoes = _series({ON: sem_negocio, PN: [("100", 5.0)] * 8})
    com, motivo = historico.data_com(series[ON], pregoes, "2016-04-27", Fraction(11, 10))
    assert com is None and motivo.startswith("papel sem negócio")
    # a janela para na aprovação seguinte do mesmo tipo (não pega o evento do ano seguinte)
    series, pregoes = _series({ON: duas})
    assert historico.data_com(series[ON], pregoes, "2016-04-27", Fraction(11, 10),
                              ate="2016-05-01") == ("2016-04-28", "")  # fmt: skip


def _fre(classe: str, antes: int, depois: int, tipo: str = "BONIFICACAO",
         aprovado: str = "2016-04-27") -> LinhaFre:  # fmt: skip
    return LinhaFre(CNPJ, aprovado, tipo, classe, antes, depois, "65000")


def test_eventos_fre_com_especie_irma_e_unit() -> None:
    on = [("100", 10.0), ("101", 9.1), ("101", 9.1), ("101", 9.1), ("101", 9.1),
          ("101", 9.1), ("101", 9.1), ("101", 9.1)]  # fmt: skip
    pn = [("200", 5.0), ("201", 5.0), ("201", 5.0), ("201", 5.0), ("201", 5.0),
          ("202", 4.9), ("202", 4.9), ("202", 4.9)]  # fmt: skip  # subiu no dia: não confirma
    unit = [("300", 30.0), ("301", 27.3)] + [("301", 27.3)] * 6
    series, pregoes = _series({ON: on, PN: pn, UNIT: unit})
    eventos, revisao = historico.eventos_fre(
        [_fre("NOR", 1000, 1100), _fre("NPR", 2000, 2200), _fre("NPA", 10, 11)],
        {CNPJ: "ITSA"}, series, pregoes)  # fmt: skip
    assert revisao == []  # NPA sem negócio: nada a ajustar
    assert sorted((e.isin, e.fator, e.data_com, e.origem) for e in eventos) == [
        (ON, "10", "2016-04-27", "cvm_fre"),
        (PN, "10", "2016-04-27", "cvm_fre"),  # a data com da ON, com mudança no mesmo pregão
        (UNIT, "10", "2016-04-27", "cvm_fre"),  # a unit caiu como o fator manda
    ]
    assert all(evento_b3_valido(e.linha()) for e in eventos)


def test_grupamento_e_desdobramento_no_mesmo_ato_vao_para_revisao() -> None:
    series, pregoes = _series({ON: [("100", 10.0), ("101", 10.0)] + [("101", 10.0)] * 6})
    eventos, revisao = historico.eventos_fre(
        [_fre("NOR", 100, 1, "GRUPAMENTO"), _fre("NOR", 1, 100, "DESDOBRAMENTO")],
        {CNPJ: "ITSA"}, series, pregoes)  # fmt: skip
    assert eventos == [] and len(revisao) == 2
    assert all("mesmo ato" in r.motivo for r in revisao)
    ruido, revisao = historico.eventos_fre([_fre("NOR", 5000000, 5491000)], {CNPJ: "ITSA"},
                                          series, pregoes)  # fmt: skip
    assert ruido == [] and [r.motivo for r in revisao] == ["fator não é uma fração simples"]
    # a B3 já traz o evento (fator exato perto da razão do FRE): sai da revisão
    b3 = {"isin": ON, "tipo": "BONIFICACAO", "fator": "9.82", "data_com": "2016-05-02"}
    assert historico.cobertas(revisao, [b3]) == []


def test_bonificacao_e_desdobramento_do_mesmo_ato_em_versoes_do_fre_vao_para_revisao() -> None:
    # RCSL 2010: uma versão do FRE diz bonificação de 300%, outra desdobramento de 300%; as duas
    # linhas multiplicariam a posição por 16 (a cotação mostra 4)
    on = [("100", 8.0), ("101", 2.0)] + [("101", 2.0)] * 6
    series, pregoes = _series({ON: on})
    eventos, revisao = historico.eventos_fre(
        [_fre("NOR", 100, 400), _fre("NOR", 100, 400, "DESDOBRAMENTO")],
        {CNPJ: "ITSA"}, series, pregoes)  # fmt: skip
    assert eventos == [] and len(revisao) == 2
    assert all("tipos diferentes" in r.motivo for r in revisao)


def _ev(fator: str, data_com: str, origem: str = "cvm_fre",
        tipo: str = "BONIFICACAO") -> EventoB3:  # fmt: skip
    return EventoB3(emissor="ITSA", isin=ON, tipo=tipo, fator=fator, data_com=data_com,
                    aprovado_em="2016-04-27", isin_emitido=ON, origem=origem,
                    referencia="FRE 65000")  # fmt: skip


def test_juntar_bonificacao_de_uma_fonte_e_desdobramento_de_outra_e_divergencia() -> None:
    # CGRA 2008: a B3 diz desdobramento de 400%, o FRE bonificação de 400% no mesmo dia
    b3 = _ev("400", "2008-09-24", "b3", "DESDOBRAMENTO").linha()
    resumo = Resumo()
    linhas = gerar.juntar_eventos([b3], [_ev("400", "2008-09-24")], resumo)
    assert [(r["tipo"], r["origem"]) for r in linhas] == [("DESDOBRAMENTO", "b3")]
    assert len(resumo.divergencias) == 1


def test_juntar_soma_origem_e_divergencia_nunca_resolvida_por_palpite() -> None:
    anterior = {**_ev("10", "2016-04-29", "b3").linha(), "origem": "", "referencia": ""}
    resumo = Resumo()
    linhas = gerar.juntar_eventos([anterior], [
        _ev("10", "2016-04-29"),  # mesmo evento: soma a origem
        _ev("9.090909091", "2016-04-29"),  # outro fator na mesma data com: divergência
        _ev("10", "2016-05-03"),  # mesmo fator em data com próxima: divergência
        _ev("10", "2017-04-28"),  # um ano depois: evento novo
    ], resumo)  # fmt: skip
    assert [(r["data_com"], r["fator"], r["origem"], r["referencia"]) for r in linhas] == [
        ("2016-04-29", "10", "b3+cvm_fre", "FRE 65000"),
        ("2017-04-28", "10", "cvm_fre", "FRE 65000"),
    ]
    assert len(resumo.divergencias) == 2 and resumo.eventos_novos == [
        "2017-04-28 ITSA BONIFICACAO 10% (cvm_fre)"
    ]
    assert not resumo.vazio
    # dízima arredondada diferente nas duas fontes é o mesmo fator
    resumo = Resumo()
    b3 = {**_ev("33.333333333", "2016-04-29", "b3").linha()}
    linhas = gerar.juntar_eventos([b3], [_ev("33.333333334", "2016-04-29")], resumo)
    assert [r["origem"] for r in linhas] == ["b3+cvm_fre"] and not resumo.divergencias


def test_eventos_curados_dos_avisos_da_cvm() -> None:
    curados = gerar.eventos_curados()
    assert len(curados) >= 30
    for e in curados:
        assert evento_b3_valido(e.linha()), e
        assert e.origem == "cvm_ipe" and e.referencia.startswith("IPE ")
        assert e.tipo in gerar.TIPOS_COM_FATOR
        # só a bonificação pode vir em outra espécie (ITSA 2018: PN para quem tem ON e PN)
        assert e.isin == e.isin_emitido or e.tipo == "BONIFICACAO"
        assert e.aprovado_em <= e.data_com
    # ITSA, nov/2023: 5% com data-base 27/11/2023 e o custo atribuído do fato relevante
    itsa = [e for e in curados if e.emissor == "ITSA" and e.data_com == "2023-11-27"]
    assert {(e.isin, e.fator, e.custo_atribuido) for e in itsa} == {
        (ON, "5", "17.917246"),
        (PN, "5", "17.917246"),
    }


def test_historico_soma_curados_e_grava_data_do_pacote(tmp_path: Path) -> None:
    (tmp_path / "eventos_b3.csv").write_text(",".join(gerar.COLUNAS_EVENTOS) + "\n", "utf-8")
    dia = date(2026, 10, 28)
    arquivos, resumo = gerar.incorporar_eventos(tmp_path, [_ev("10", "2011-04-29")], dia)
    assert '"data": "2026-10-28"' in arquivos["pacote.json"]
    assert len(resumo.eventos_novos) == len(gerar.eventos_curados()) + 1
    assert gerar.gravar(tmp_path, arquivos) == ["eventos_b3.csv"]
    assert timedelta(days=270) == historico.JANELA
