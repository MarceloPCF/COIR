"""Ativo isolado: um evento que falta não para a apuração inteira (decisões do usuário de
03/11/2026, parecer do especialista). O ativo sai do cálculo desde o problema; o valor das vendas
dele continua no limite de isenção; os meses dele desde a última posição zerada ficam
incompletos e os seguintes da mesma categoria, dependentes. Venda que passa de comprado a
vendido e assim fica no fim do mês só gera aviso."""

from datetime import date

import pytest

from coir.domain.fiscal import AtivoIsolado, Categoria, TipoAtivo, TipoEvento
from coir.engine.movimentos import carteira, movimentar
from coir.erros import MotorError
from coir.servicos.darf import pendencia_do_mes

from .apoio import REFS, D, Livro, categoria, mes

VENC = date(2025, 3, 21)


def _livro() -> Livro:
    """Call de ITSA4 lançada sobre 100 ações; a bonificação que ajustaria a série não foi lançada
    e o exercício vem com 110: o motor não fecha a série no exercício."""
    livro = Livro()
    livro.nota(date(2025, 1, 10), livro.op("C", "ITSA4", 100, "10"),
               livro.op("V", "ITSAC100", 100, "0.5", mercado="CALL", vencimento=VENC),
               livro.op("C", "PETR4", 1000, "20"), livro.op("C", "BBAS3", 2000, "10"))  # fmt: skip
    livro.nota(date(2025, 2, 12), livro.op("C", "ITSA4", 100, "10"))
    livro.nota(VENC, livro.op("V", "ITSA4", 110, "10", mercado="EXERC CALL", opcao="ITSAC100"))
    # abril: 18.000 de PETR4 + 3.000 de ITSA4 (isolada) = 21.000 de vendas: sem isenção
    livro.nota(date(2025, 4, 10), livro.op("V", "PETR4", 600, "30"),
               livro.op("V", "ITSA4", 100, "30"))  # fmt: skip
    livro.nota(date(2025, 5, 12), livro.op("V", "BBAS3", 2000, "12"))  # ganho de 4.000, tributável
    return livro


def test_sem_a_lista_de_isolados_o_motor_continua_parando() -> None:
    livro = _livro()
    with pytest.raises(MotorError, match="posição vendida na opção") as erro:
        movimentar(livro.notas, livro.eventos, REFS.parametros, date(2025, 5, 31))
    assert erro.value.data == VENC


def test_ativo_isolado_sai_do_calculo_e_o_resto_segue() -> None:
    ap = _livro().apurar(date(2025, 5, 31))
    por_ticker = {i.ticker: i for i in ap.isolados}
    assert set(por_ticker) == {"ITSA4", "ITSAC100"}
    itsa = por_ticker["ITSA4"]
    assert itsa.desde == VENC and "posição vendida na opção" in itsa.motivo
    # desde a última posição zerada (janeiro), não só a partir do exercício
    assert itsa.meses == ((2025, 1), (2025, 2), (2025, 3), (2025, 4))
    assert por_ticker["ITSAC100"].meses == ((2025, 1), (2025, 3))
    # PETR4 e BBAS3 apurados normalmente
    assert categoria(ap, 2025, 5, Categoria.COMUM).resultado == D("4000")
    assert any("ITSA4 ficou fora do cálculo" in a for a in ap.avisos)


def test_venda_do_ativo_isolado_conta_no_limite_de_isencao() -> None:
    ap = _livro().apurar(date(2025, 5, 31))
    abril = categoria(ap, 2025, 4, Categoria.COMUM)
    assert abril.vendas_acoes == D("21000")
    # o ganho de PETR4 (6.000) é tributável: sem a venda de ITSA4 seria isento
    assert (abril.isento, abril.resultado, abril.imposto) == (0, D("6000"), D("900"))


def test_meses_incompletos_e_dependentes() -> None:
    ap = _livro().apurar(date(2025, 5, 31))
    abril, maio = mes(ap, 2025, 4), mes(ap, 2025, 5)
    assert abril.incompletos == ("ITSA4",) and abril.provisorio
    assert maio.incompletos == () and set(maio.dependentes) == {"ITSA4", "ITSAC100"}
    assert "Valor provisório" in pendencia_do_mes(abril)
    assert "isolado(s) em mês anterior" in pendencia_do_mes(maio)
    assert pendencia_do_mes(mes(ap, 2025, 1)).startswith("Valor provisório")


def test_carteira_sem_o_ativo_isolado() -> None:
    livro = _livro()
    isolados: list[AtivoIsolado] = []
    posicoes = carteira(livro.notas, livro.eventos, REFS.parametros, date(2025, 3, 31),
                        isolados=isolados)  # fmt: skip
    assert {p.ticker for p in posicoes} == {"PETR4", "BBAS3"}
    assert {i.ticker for i in isolados} == {"ITSA4", "ITSAC100"}


def test_venda_que_passa_de_comprado_a_vendido_avisa() -> None:
    livro = Livro()
    livro.nota(date(2025, 1, 10), livro.op("C", "MGLU3", 100, "10"))
    # desdobramento de 1 para 3 não lançado: o usuário vende as 300 que tem
    livro.nota(date(2025, 2, 10), livro.op("V", "MGLU3", 300, "4"))
    ap = livro.apurar(date(2025, 2, 28))
    assert any("2025-02-10: venda de 300 MGLU3 maior que a posição (100), que ficou vendida no "
               "fim do mês (-200)" in a for a in ap.avisos)  # fmt: skip
    assert not ap.isolados  # só avisa: o cálculo não muda


def test_cada_aviso_fica_com_o_titular_dele() -> None:
    """Decisão do usuário de 05/11/2026: o sino, o Painel e a Apuração mostram só os avisos do
    titular escolhido. A posição de um titular não cobre a venda do outro, e o ativo isolado de um
    não aparece para o outro."""
    livro = _livro()
    livro.titular = 2
    livro.nota(date(2025, 1, 10), livro.op("C", "MGLU3", 100, "10"))
    livro.nota(date(2025, 2, 10), livro.op("V", "MGLU3", 300, "4"))
    ap = livro.apurar(date(2025, 5, 31))

    [do_2] = ap.avisos_de(2)
    assert "venda de 300 MGLU3 maior que a posição (100)" in do_2
    do_1 = ap.avisos_de(1)
    assert any("ITSA4 ficou fora do cálculo" in a for a in do_1)
    assert not any("MGLU3" in a for a in do_1)
    assert ap.avisos_de(3) == ()
    assert set(ap.avisos) == {*do_1, do_2}


def test_vendida_e_recomprada_no_mesmo_mes_nao_avisa() -> None:
    """Decisão do usuário de 03/11/2026: só avisa se continuar vendida no fim do mês."""
    livro = Livro()
    livro.nota(date(2025, 1, 10), livro.op("C", "CSNA3", 30, "10"))
    livro.nota(date(2025, 2, 10), livro.op("V", "CSNA3", 300, "10"))
    livro.nota(date(2025, 2, 20), livro.op("C", "CSNA3", 270, "9"))
    assert not any("maior que a posição" in a for a in livro.apurar(date(2025, 2, 28)).avisos)


def test_venda_a_descoberto_a_partir_de_zero_nao_avisa() -> None:
    livro = Livro()
    livro.nota(date(2025, 2, 10), livro.op("V", "MGLU3", 300, "4"))
    livro.nota(date(2025, 3, 10), livro.op("C", "MGLU3", 300, "3"))
    assert not any("maior que a posição" in a for a in livro.apurar(date(2025, 3, 31)).avisos)


def test_day_trade_que_compensa_a_venda_nao_avisa() -> None:
    livro = Livro()
    livro.nota(date(2025, 1, 10), livro.op("C", "MGLU3", 100, "10"))
    livro.nota(date(2025, 2, 10), livro.op("V", "MGLU3", 200, "11"),
               livro.op("C", "MGLU3", 100, "10"))  # fmt: skip
    assert not any("maior que a posição" in a for a in livro.apurar(date(2025, 2, 28)).avisos)


def test_serie_puxada_herda_o_motivo_de_quem_a_puxou() -> None:
    """Dois problemas no mesmo titular: ITSA4 (exercício em março) e depois LREN3 (exercício de
    110 sobre 100 em abril). A série LRENF12, exercida em maio sobre LREN3 já isolada, sai junto e
    o motivo dela é o de LREN3, não o primeiro problema do titular (simulação da 3.1.0 com as
    notas reais: as opções do LREN mostravam o motivo do ITSA)."""
    livro = _livro()
    livro.nota(date(2025, 1, 10), livro.op("C", "LREN3", 100, "15"),
               livro.op("C", "LRENC10", 100, "0.5", mercado="CALL",
                        vencimento=date(2025, 4, 17)))  # fmt: skip
    livro.nota(date(2025, 4, 17), livro.op("C", "LREN3", 110, "10", mercado="EXERC CALL",
                                           opcao="LRENC10"))  # fmt: skip
    livro.nota(date(2025, 5, 2), livro.op("C", "LRENF12", 100, "0.4", mercado="CALL",
                                          vencimento=date(2025, 5, 16)))  # fmt: skip
    livro.nota(date(2025, 5, 16), livro.op("C", "LREN3", 100, "12", mercado="EXERC CALL",
                                           opcao="LRENF12"))  # fmt: skip
    por_ticker = {i.ticker: i for i in livro.apurar(date(2025, 5, 31)).isolados}
    assert {"ITSA4", "LREN3", "LRENC10", "LRENF12"} <= set(por_ticker)
    assert "LRENC10" in por_ticker["LREN3"].motivo
    assert por_ticker["LRENF12"].motivo == por_ticker["LREN3"].motivo
    assert "ITSAC100" in por_ticker["ITSA4"].motivo


def _com_ativo_sem_tipo() -> Livro:
    """Um direito de subscrição (sem tipo nas tabelas) e ações que seguem apuradas."""
    livro = Livro()
    livro.nota(date(2025, 1, 10), livro.op("C", "ZZZZ12", 10, "5", tipo=TipoAtivo.SEM_TIPO),
               livro.op("C", "PETR4", 1000, "20"))  # fmt: skip
    livro.nota(date(2025, 2, 10), livro.op("V", "ZZZZ12", 10, "2000", tipo=TipoAtivo.SEM_TIPO))
    livro.nota(date(2025, 3, 10), livro.op("V", "PETR4", 1000, "30"))  # 30.000: ganho tributável
    return livro


def test_ativo_sem_tipo_fica_isolado_e_nao_para_a_apuracao() -> None:
    ap = _com_ativo_sem_tipo().apurar(date(2025, 3, 31))
    [isolado] = ap.isolados
    assert (isolado.ticker, isolado.desde) == ("ZZZZ12", date(2025, 1, 10))
    assert "não tem tipo de ativo" in isolado.motivo
    assert "data/tipos_ativo" not in isolado.motivo  # quem lê é o usuário
    assert isolado.meses == ((2025, 1), (2025, 2))
    assert categoria(ap, 2025, 3, Categoria.COMUM).resultado == D("10000")
    assert mes(ap, 2025, 2).incompletos == ("ZZZZ12",)


def test_venda_do_ativo_sem_tipo_nao_conta_como_venda_de_acao() -> None:
    """O que não é ação não entra no limite (SC Cosit 145/2021): 20 mil de ações seguem isentos."""
    livro = Livro()
    livro.nota(date(2025, 1, 10), livro.op("C", "ZZZZ12", 10, "5", tipo=TipoAtivo.SEM_TIPO),
               livro.op("C", "PETR4", 1000, "10"))  # fmt: skip
    livro.nota(date(2025, 2, 10), livro.op("V", "ZZZZ12", 10, "2000", tipo=TipoAtivo.SEM_TIPO),
               livro.op("V", "PETR4", 1000, "19"))  # fmt: skip
    fevereiro = categoria(livro.apurar(date(2025, 2, 28)), 2025, 2, Categoria.COMUM)
    assert fevereiro.vendas_acoes == D("19000")
    assert (fevereiro.isento, fevereiro.imposto) == (D("9000"), 0)


def test_evento_de_ativo_sem_tipo_tambem_isola() -> None:
    livro = Livro()
    livro.evento(date(2024, 12, 31), "ZZZZ12", TipoEvento.CUSTODIA_INICIAL, cv="C", quantidade="10",
                 valor="50", tipo_ativo=TipoAtivo.SEM_TIPO)  # fmt: skip
    livro.nota(date(2025, 1, 10), livro.op("C", "PETR4", 10, "20"))
    assert [i.ticker for i in livro.apurar(date(2025, 1, 31)).isolados] == ["ZZZZ12"]
