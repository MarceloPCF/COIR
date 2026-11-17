"""Data da liquidação do termo pela movimentação da B3 (linha "Liquidação Termo"): antecipada,
parcial, no vencimento e sem a movimentação do período (vencimento com aviso). A regra fiscal
(alienação na liquidação, no vencimento ou antecipada: IN RFB 1.585/2015 art. 59 §2º IV) está
nos casos de tests/fiscal_cases/test_termo_liquidacao_antecipada.py."""

import sqlite3
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest

from coir import referencias
from coir.domain.fiscal import Visao
from coir.servicos.apuracao import apurar_banco, carteira_banco, termos_banco
from coir.servicos.b3 import importar_planilhas
from coir.servicos.termo import (
    Contrato,
    LinhaLiquidacao,
    codigos_do_produto,
    e_liquidacao_termo,
    liquidacoes,
)
from tests.unit.test_importador_b3 import CAB_MOV, planilha
from tools.banco_demo import criar

D = Decimal
AGORA = datetime(2026, 10, 19, 12, 0, 0)
REFS, FISCAIS = referencias.carregar(), referencias.carregar_fiscais()
AGORA_B3 = "AGORA CORRETORA DE TITULOS E VALORES MOBILIARIOS S/A"
PRODUTO = "Termo Acoes - GGBR4T - GERDAU S.A."


def _contrato(chave: int = 1, quantidade: str = "300", cv: str = "C", **troca: object) -> Contrato:
    dados: dict[str, object] = {
        "chave": chave, "titular": 1, "corretora": "AGORA", "ativo": 7, "ticker": "GGBR4",
        "cv": cv, "pregao": date(2026, 3, 2), "vencimento": date(2026, 5, 4),
        "quantidade": D(quantidade),
    }  # fmt: skip
    dados.update(troca)
    return Contrato(**dados)  # type: ignore[arg-type]


def _linha(dia: date, quantidade: str, entrada: bool = True, ativo: int = 7) -> LinhaLiquidacao:
    return LinhaLiquidacao(titular=1, corretora="AGORA", ativo=ativo, entrada=entrada, data=dia,
                           quantidade=D(quantidade))  # fmt: skip


PERIODO = {1: (date(2026, 1, 2), date(2026, 6, 30))}


def test_rotulo_e_codigo_da_linha_de_termo() -> None:
    assert e_liquidacao_termo("Liquidação Termo") and e_liquidacao_termo(" LIQUIDACAO  TERMO ")
    assert not e_liquidacao_termo("Transferência - Liquidação")
    assert codigos_do_produto("Termo Acoes - PETR4T - PETROBRAS", "TERMO") == ["PETR4", "TERMO"]
    assert codigos_do_produto("Termo de Ação PETR4T", "PETR4T") == ["PETR4", "PETR4T"]
    assert codigos_do_produto("TAEE11T - TAESA", "TAEE11T") == ["TAEE11", "TAEE11T"]
    assert codigos_do_produto("PETR4 - PETROBRAS", "PETR4") == ["PETR4"]


def test_liquidacao_antecipada_total() -> None:
    liq = liquidacoes([_contrato()], [_linha(date(2026, 3, 16), "300")], PERIODO)[1]
    assert liq.partes == ((date(2026, 3, 16), D(300)),) and liq.aviso == ""


def test_liquidacao_no_vencimento_sem_aviso() -> None:
    liq = liquidacoes([_contrato()], [_linha(date(2026, 5, 4), "300")], PERIODO)[1]
    assert liq.partes == ((date(2026, 5, 4), D(300)),) and liq.aviso == ""


def test_antecipada_parcial_e_o_resto_no_vencimento() -> None:
    linhas = [_linha(date(2026, 3, 16), "100"), _linha(date(2026, 5, 4), "200")]
    liq = liquidacoes([_contrato()], linhas, PERIODO)[1]
    assert liq.partes == ((date(2026, 3, 16), D(100)), (date(2026, 5, 4), D(200)))
    assert liq.aviso == ""


def test_venda_a_termo_casa_com_o_debito() -> None:
    linhas = [_linha(date(2026, 3, 16), "300", entrada=True),
              _linha(date(2026, 4, 1), "300", entrada=False)]  # fmt: skip
    liq = liquidacoes([_contrato(cv="V")], linhas, PERIODO)[1]
    assert liq.partes == ((date(2026, 4, 1), D(300)),)


def test_linha_fora_do_contrato_ou_de_outro_ativo_nao_conta() -> None:
    linhas = [_linha(date(2026, 3, 2), "300"),  # no próprio pregão
              _linha(date(2026, 5, 5), "300"),  # depois do vencimento
              _linha(date(2026, 3, 16), "300", ativo=8)]  # fmt: skip
    liq = liquidacoes([_contrato()], linhas, PERIODO)[1]
    assert liq.partes == ((date(2026, 5, 4), D(300)),)
    assert "não mostra a liquidação de 300 ações da compra a termo de 02/03/2026" in liq.aviso


def test_dois_contratos_do_mesmo_ativo_cada_um_com_a_sua_linha() -> None:
    """A linha com a quantidade exata do contrato vem antes da ordem das datas."""
    contratos = [_contrato(1, "300"), _contrato(2, "100", pregao=date(2026, 3, 3),
                                                vencimento=date(2026, 5, 5))]  # fmt: skip
    linhas = [_linha(date(2026, 3, 10), "100"), _linha(date(2026, 3, 20), "300")]
    liq = liquidacoes(contratos, linhas, PERIODO)
    assert liq[1].partes == ((date(2026, 3, 20), D(300)),)
    assert liq[2].partes == ((date(2026, 3, 10), D(100)),)


def test_sem_movimentacao_a_venda_a_termo_sempre_avisa() -> None:
    """Quem antecipa é o comprador: o vendedor não tem como saber sem a B3 (parecer de
    17/11/2026)."""
    sem = liquidacoes([_contrato(cv="V")], [], {})[1]
    curta = liquidacoes([_contrato(cv="V")], [], {1: (date(2026, 3, 1), date(2026, 4, 30))})[1]
    for liq in (sem, curta):
        assert liq.partes == ((date(2026, 5, 4), D(300)),)
        assert liq.aviso.startswith("Termo: o COIR usou o vencimento, 04/05/2026")
        assert "venda a termo" in liq.aviso and "importe a movimentação da B3" in liq.aviso


def test_sem_movimentacao_a_compra_a_termo_so_avisa_com_venda_na_janela_ou_outro_ano() -> None:
    """Compra a termo: venda à vista do ativo entre o pregão e o vencimento pode ter exigido a
    liquidação antecipada (entrega sem ações em custódia); vencimento em outro ano muda o ano da
    apuração. Fora disso, sem aviso (parecer de 17/11/2026)."""
    assert liquidacoes([_contrato()], [], {})[1].aviso == ""
    fora = liquidacoes([_contrato()], [], {}, {(1, 7): [date(2026, 5, 4)]})[1]
    assert fora.aviso == ""  # venda no dia do vencimento: já com as ações do termo
    vendeu = liquidacoes([_contrato()], [], {}, {(1, 7): [date(2026, 4, 9)]})[1]
    assert vendeu.partes == ((date(2026, 5, 4), D(300)),)
    assert vendeu.aviso.startswith("Termo: houve venda à vista de GGBR4 em 09/04/2026")
    virada = liquidacoes([_contrato(pregao=date(2025, 12, 1), vencimento=date(2026, 3, 2))],
                         [], {})[1]  # fmt: skip
    assert virada.aviso.startswith("Termo: o COIR usou o vencimento, 02/03/2026")


# ------------------------------------------------------------------- com o banco (demonstração)
@pytest.fixture
def termo(tmp_path: Path) -> Path:
    """Banco de demonstração com a compra de 300 GGBR4 a 62,20 de 18/04/2008 (Agora, nota só
    com ela) transformada em compra a termo de 60 dias: vencimento em 17/06/2008."""
    banco = criar(tmp_path / "demo")
    con = sqlite3.connect(banco)
    with con:
        con.execute("update operacao set mercado='TERMO', prazo_termo_dias=60 where id=1")
    con.close()
    return banco


def _movimentacao(banco: Path, *linhas: list[object]) -> None:
    arquivo = planilha(banco.parent / "movimentacao-2008.xlsx", {"Movimentação": [
        CAB_MOV,
        ["Credito", "01/04/2008", "Dividendo", "BBAS3 - BANCO DO BRASIL S.A.", AGORA_B3, 1, 1, 1],
        *linhas,
        ["Credito", "01/12/2008", "Dividendo", "BBAS3 - BANCO DO BRASIL S.A.", AGORA_B3, 1, 1, 1],
    ]})  # fmt: skip
    importar_planilhas(banco, [arquivo], 1, AGORA)


def _ggbr4(banco: Path, dia: date) -> Decimal:
    pos, _ = carteira_banco(banco, REFS, FISCAIS, dia, Visao.TITULAR)
    return next((p.quantidade for p in pos if p.ticker == "GGBR4"), D(0))


def _avisos_termo(banco: Path) -> list[str]:
    ap = apurar_banco(banco, REFS, FISCAIS, date(2008, 12, 31))
    return [a for a in ap.avisos if a.startswith("Termo:")]


def test_banco_sem_movimentacao_liquida_no_vencimento(termo: Path) -> None:
    """Compra a termo sem venda do ativo no prazo do contrato: vencimento, sem aviso."""
    assert _ggbr4(termo, date(2008, 6, 16)) == 0
    assert _ggbr4(termo, date(2008, 6, 17)) == 300
    assert _avisos_termo(termo) == []


def test_banco_venda_a_termo_sem_movimentacao_avisa(termo: Path) -> None:
    """A mesma operação como venda a termo descoberta (sem ações para entregar o motor para):
    aqui só o aviso, pelas entradas do serviço."""
    from coir.db.sessao import conectar
    from coir.servicos.apuracao import entradas

    con = sqlite3.connect(termo)
    with con:
        con.execute("update operacao set cv='V' where id=1")
    con.close()
    with conectar(termo) as fabrica, fabrica() as s:
        notas, _ = entradas(s, REFS, FISCAIS)
    [op] = [o for n in notas for o in n.operacoes if o.id == 1]
    assert op.liquidacao_termo == date(2008, 6, 17)
    assert "17/06/2008" in op.aviso and "GGBR4" in op.aviso and "AGORA" in op.aviso


def test_banco_com_liquidacao_antecipada_pela_b3(termo: Path) -> None:
    _movimentacao(termo, ["Credito", "05/05/2008", "Liquidação Termo", PRODUTO, AGORA_B3, 300,
                          62.2, 18660])  # fmt: skip
    assert _ggbr4(termo, date(2008, 5, 2)) == 0
    assert _ggbr4(termo, date(2008, 5, 5)) == 300
    assert _avisos_termo(termo) == []
    assert termos_banco(termo, REFS, FISCAIS, date(2008, 5, 2))[0].liquidacao == date(2008, 5, 5)


def test_banco_com_liquidacao_antecipada_parcial(termo: Path) -> None:
    _movimentacao(termo, ["Credito", "05/05/2008", "Liquidação Termo", PRODUTO, AGORA_B3, 100,
                          62.2, 6220])  # fmt: skip
    assert _ggbr4(termo, date(2008, 5, 5)) == 100
    assert _ggbr4(termo, date(2008, 6, 17)) == 300
    pos, _ = carteira_banco(termo, REFS, FISCAIS, date(2008, 6, 17), Visao.TITULAR)
    sem_termo = next(p for p in pos if p.ticker == "GGBR4")
    pos_cheia, _ = carteira_banco(termo, REFS, FISCAIS, date(2008, 5, 5), Visao.TITULAR)
    parcial = next(p for p in pos_cheia if p.ticker == "GGBR4")
    assert parcial.custo * 3 == pytest.approx(sem_termo.custo, abs=D("0.02"))
    [aviso] = _avisos_termo(termo)
    assert "não mostra a liquidação de 200 ações" in aviso
