"""Planilha externa de operações finalizadas (sintética) e comparação com o motor."""

from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import Workbook

from coir import referencias
from coir.domain.fiscal import Categoria, Grupo, Movimento, TipoAtivo
from coir.erros import GabaritoError
from coir.gabarito.externo import comparar_externo, explicar, ler_planilha_externa, ticker_atual
from tests.apoio import cpf_de_teste

D = Decimal
CABECALHO = ["Ativo", "Quantidade", "Corretora", "C/V", "Custo Médio", "C/V", "Custo Final",
             "Data", "Lucratividade", "Observação", "LAIR"]  # fmt: skip


def _planilha(caminho: Path, total_jan: float = 150.0) -> Path:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.append([None, "Operações Finalizadas"])
    ws.append([None, f"01/01/2020 - 31/03/2020 • Titular de Teste • {cpf_de_teste()} • Emitido"])
    ws.append([None, "2020 - Janeiro - Bovespa"])
    ws.append([None, *CABECALHO])
    ws.append([None, "AAAA3", 100, "XP", "C", 10.0, "V", 11.0, "20/01/2020", "10,00%", None, 100])
    ws.append([None, "BBBB3", 100, "BANCO BTG PACTUAL S/A", "C", 10.0, "V", 10.5, "21/01/2020",
               "5,00%", "D", 50])  # fmt: skip
    ws.append([None, None, None, None, None, None, None, None, None, None, "Total:", total_jan])
    ws.append([None, "2020 - Fevereiro - BMF"])
    ws.append([None, *CABECALHO])
    ws.append([None, "WINH20", 1, "XP", "C", 100.0, "V", 110.0, "10/02/2020", "1%", None, -30])
    ws.append([None, None, None, None, None, None, None, None, None, None, "Total:", -30])
    ws.append([None, None, None, None, None, None, None, None, None, None, "Total:", 120])
    wb.save(caminho)
    return caminho


def _mov(dia: date, ticker: str, resultado: str, corretora: str = "XP",
         ativo: int | None = 1) -> Movimento:  # fmt: skip
    return Movimento(
        data=dia, titular=1, conta=1, corretora=corretora, ativo=ativo, ticker=ticker,
        origem="op:1", regra="venda", categoria=Categoria.COMUM, grupo=Grupo.A_VISTA,
        tipo=TipoAtivo.ACAO, cv="V", quantidade=D(100), valor=D(0), custos=D(0),
        resultado=D(resultado), quantidade_apos=D(0), custo_apos=D(0),
    )  # fmt: skip


def test_le_vendas_e_ignora_o_cabecalho_com_nome_e_cpf(tmp_path: Path) -> None:
    vendas = ler_planilha_externa(_planilha(tmp_path / "s.xlsx"))
    assert [(v.ticker, v.corretora, v.mercado, v.lair, v.day_trade) for v in vendas] == [
        ("AAAA3", "XP", "Bovespa", D(100), False), ("BBBB3", "BTG", "Bovespa", D(50), True),
        ("WINH20", "XP", "BMF", D(-30), False),
    ]  # fmt: skip
    assert cpf_de_teste() not in repr(vendas)


def test_total_da_secao_que_nao_fecha_falha_alto(tmp_path: Path) -> None:
    with pytest.raises(GabaritoError, match="Janeiro"):
        ler_planilha_externa(_planilha(tmp_path / "s.xlsx", total_jan=151.0))


def test_comparacao_classifica_e_aceita_explicacoes(tmp_path: Path) -> None:
    """Confere; mesmo dia/corretora/valor com outro código; outro dia do mesmo mês; explicada."""
    vendas = ler_planilha_externa(_planilha(tmp_path / "s.xlsx"))
    movs = [
        _mov(date(2020, 1, 20), "AAAA3", "100"),
        _mov(date(2020, 1, 21), "CCCC3", "50", corretora="BTG"),  # código diferente
        _mov(date(2020, 2, 10), "WINH20", "-20", ativo=None),
        _mov(date(2020, 2, 12), "WINH20", "-10", ativo=None),  # resto em outro dia do mês
    ]
    linhas, meses = comparar_externo(vendas, movs, lambda t: t, D("0.01"), D(1))
    classes = {(x.data.day, x.ticker): x.classe for x in linhas}
    assert classes == {
        (20, "AAAA3"): "confere", (21, "BBBB3"): "ticker com outro código",
        (21, "CCCC3"): "ticker com outro código", (10, "WINH20"): "outro dia do mesmo mês",
        (12, "WINH20"): "outro dia do mesmo mês",
    }  # fmt: skip
    assert meses == []  # os meses fecham
    chave = (date(2020, 1, 21), "BTG", "BBBB3")
    explicadas = explicar(linhas, {chave: "série renomeada"})
    assert [x.motivo for x in explicadas if x.classe == "explicada"] == ["série renomeada"]


def test_marco_desde_ignora_os_dois_lados_antes_dele(tmp_path: Path) -> None:
    """Antes do marco (ex.: carteira zerada) a planilha pode ter PM inicial lançado à mão."""
    vendas = ler_planilha_externa(_planilha(tmp_path / "s.xlsx"))
    movs = [
        _mov(date(2020, 1, 20), "AAAA3", "999"),  # diverge, mas antes do marco
        _mov(date(2020, 2, 10), "WINH20", "-30", ativo=None),
    ]
    linhas, meses = comparar_externo(vendas, movs, lambda t: t, D("0.01"), D(1), date(2020, 2, 1))
    assert [(x.data, x.ticker, x.classe) for x in linhas] == [
        (date(2020, 2, 10), "WINH20", "confere")
    ]
    assert meses == []


def test_ticker_atual_segue_a_cadeia_de_trocas_e_a_serie() -> None:
    atual = ticker_atual(referencias.carregar())
    assert atual("TRPL4") == "ISAE4" and atual("EMBRE790") == "EMBJE790"
    assert atual("PETR4") == "PETR4"
