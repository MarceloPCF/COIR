"""F8: data do crédito do JCP pelo pacote público da B3 e pela correção do usuário (dados fictícios;
o JCP de ITAUSA é do pacote público)."""

from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
from sqlalchemy import select

from coir.db.modelos import MovimentacaoB3
from coir.db.sessao import conectar
from coir.domain.b3 import LinhaMovimentacao
from coir.domain.proventos import LinhaB3
from coir.erros import ImportacaoError
from coir.servicos.b3 import importar_planilhas
from coir.servicos.cadastros import adicionar_titular
from coir.servicos.estado import Estado
from coir.servicos.proventos import creditos_b3, informar_credito
from tests.apoio import cpf_de_teste
from tests.unit.test_importador_b3 import AGORA, CAB_MOV, XP, planilha

D = Decimal
JCP = "Juros Sobre Capital Próprio"
PUBLICOS = [
    {"emissor": "ITSA", "isin": "BRITSAACNOR1", "tipo": "JRS CAP PROPRIO", "valor": "0.0242425",
     "aprovado_em": "2024-12-16", "data_com": "2024-12-19", "pagamento": "2025-03-10"},
    {"emissor": "ITSA", "isin": "BRITSAACNPR7", "tipo": "JRS CAP PROPRIO", "valor": "0.0242425",
     "aprovado_em": "2024-12-16", "data_com": "2024-12-19", "pagamento": "2025-03-10"},
    {"emissor": "BBDC", "isin": "BRBBDCACNOR1", "tipo": "JRS CAP PROPRIO", "valor": "0.32385",
     "aprovado_em": "2025-09-30", "data_com": "2025-10-13", "pagamento": "2026-04-30"},
    {"emissor": "BBDC", "isin": "BRBBDCACNPR8", "tipo": "JRS CAP PROPRIO", "valor": "0.35624",
     "aprovado_em": "2025-11-30", "data_com": "2025-12-13", "pagamento": "2026-04-30"},
]  # fmt: skip


def _linha(codigo: str, dia: date, preco: str, i: int = 1) -> LinhaB3:
    return LinhaB3(id=i, titular=1, chave=f"k{i}", linha=LinhaMovimentacao(
        data=dia, entrada=True, movimentacao=JCP, produto=f"{codigo} - X", codigo=codigo,
        instituicao=XP, quantidade=D(1000), preco=D(preco), valor=D("20.61")))  # fmt: skip


def test_credito_pela_b3_casando_empresa_pagamento_e_valor() -> None:
    """O crédito é a mais tarde entre a aprovação e a data com (IN SRF 41/1998 art. 3º; IN RFB
    1.700/2017 art. 75 §10; decisão do usuário de 21/10/2026)."""
    itsa = _linha("ITSA4", date(2025, 3, 10), "0.02424250")
    on, pn = _linha("BBDC3", date(2026, 4, 30), "0.3238587", 2), _linha("BBDC4",
                                                                       date(2026, 4, 30),
                                                                       "0.3562445", 3)  # fmt: skip
    outro_dia = _linha("ITSA4", date(2025, 3, 11), "0.0242425", 4)
    informado = _linha("ITSA4", date(2025, 3, 10), "0.0242425", 5).model_copy(
        update={"credito": date(2024, 12, 1), "credito_origem": "informada"}
    )
    avisos: list[str] = []
    saida = creditos_b3([itsa, on, pn, outro_dia, informado], PUBLICOS, {"BBDC3": "BBDC"},
                        avisos)  # fmt: skip
    assert [(x.credito, x.credito_origem) for x in saida] == [
        (date(2024, 12, 19), "B3"),  # data com depois da aprovação: crédito na data com
        (date(2025, 10, 13), "B3"),  # ON: o valor mais próximo
        (date(2025, 12, 13), "B3"),  # PN (emissor pela raiz do ticker)
        (None, ""),  # pagamento em outra data: sem crédito
        (date(2024, 12, 1), "informada"),  # a correção do usuário vale
    ]
    # informada antes da data com da B3: vale, com aviso (decisão do usuário de 21/10/2026)
    assert avisos == ["JCP de ITSA4 pago em 10/03/2025: crédito informado em 01/12/2024, antes "
                      "da data com 19/12/2024 da B3 (confira o informe)"]  # fmt: skip


def _banco(tmp_path: Path) -> Estado:
    estado = Estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    mov = planilha(tmp_path / "movimentacao-2026-10-19.xlsx", {"Movimentação": [
        CAB_MOV,
        ["Credito", "10/03/2025", JCP, "ITSA4 - ITAUSA", XP, 1000, "0,5", "425,00"],
        ["Credito", "11/03/2025", "Dividendo", "ITSA4 - ITAUSA", XP, 1000, "0,1", "100,00"],
    ]})  # fmt: skip
    importar_planilhas(estado.banco, [mov], 1, AGORA)
    estado.titular = 1
    return estado


def test_correcao_do_usuario_muda_o_ano_do_jcp(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from coir import referencias

    monkeypatch.setattr(referencias, "proventos_b3", lambda _pasta: [])  # sem o pacote real
    estado = _banco(tmp_path)
    with conectar(estado.banco) as fabrica, fabrica() as s:
        chaves = {m.movimentacao: m.chave for m in s.scalars(select(MovimentacaoB3))}
    [jcp] = [p for p in estado.proventos()[0] if p.chave == chaves[JCP]]
    assert (jcp.ano_declaracao, jcp.credito) == (2025, None)
    informar_credito(estado.banco, chaves[JCP], date(2024, 12, 16), AGORA)
    estado = Estado(estado.banco)  # sem o cache da leitura anterior
    estado.preparar(AGORA)
    estado.titular = 1
    [jcp] = [p for p in estado.proventos()[0] if p.chave == chaves[JCP]]
    assert (jcp.ano_declaracao, jcp.credito, jcp.a_receber_em_31_12) == (
        2024, date(2024, 12, 16), True)  # fmt: skip
    assert "(informada)" in jcp.observacao
    informar_credito(estado.banco, chaves[JCP], None, AGORA)  # apagar volta à regra anterior
    estado = Estado(estado.banco)
    estado.preparar(AGORA)
    estado.titular = 1
    [jcp] = [p for p in estado.proventos()[0] if p.chave == chaves[JCP]]
    assert jcp.credito is None


def test_correcao_invalida_falha_alto(tmp_path: Path) -> None:
    estado = _banco(tmp_path)
    with conectar(estado.banco) as fabrica, fabrica() as s:
        chaves = {m.movimentacao: m.chave for m in s.scalars(select(MovimentacaoB3))}
    with pytest.raises(ImportacaoError, match="depois do pagamento"):
        informar_credito(estado.banco, chaves[JCP], date(2025, 3, 11), AGORA)
    with pytest.raises(ImportacaoError, match="só o JCP"):
        informar_credito(estado.banco, chaves["Dividendo"], date(2025, 1, 1), AGORA)
    with pytest.raises(ImportacaoError, match="não encontrada"):
        informar_credito(estado.banco, "nada", date(2025, 1, 1), AGORA)


def test_aprovacao_depois_do_pagamento_no_pacote_e_ignorada_com_aviso() -> None:
    """Dado público retificado (aprovação depois do pagamento) não derruba os proventos: a linha é
    ignorada, com aviso; só a data digitada pelo usuário falha alto (revisão da F8)."""
    ruim = [dict(PUBLICOS[0], aprovado_em="2025-04-01")]
    avisos: list[str] = []
    [x] = creditos_b3([_linha("ITSA4", date(2025, 3, 10), "0.02424250")], ruim, {}, avisos)
    assert x.credito is None
    assert avisos and "ignorado" in avisos[0]


HISTORICO = [
    {"emissor": "ITSA", "classe": "PN", "tipo": "JRS CAP PROPRIO", "valor": "0.0235295",
     "aprovado_em": "2025-02-10", "data_com": "2025-08-29"},
    {"emissor": "ITSA", "classe": "PN", "tipo": "JRS CAP PROPRIO", "valor": "0.0235295",
     "aprovado_em": "2025-02-10", "data_com": "2025-11-28"},
    {"emissor": "ITSA", "classe": "PN", "tipo": "JRS CAP PROPRIO", "valor": "0.0182",
     "aprovado_em": "2025-12-01", "data_com": "2025-12-09"},
    {"emissor": "XXXX", "classe": "ON", "tipo": "JRS CAP PROPRIO", "valor": "0.5",
     "aprovado_em": "2023-03-01", "data_com": "2023-03-10"},
    {"emissor": "XXXX", "classe": "ON", "tipo": "JRS CAP PROPRIO", "valor": "0.5",
     "aprovado_em": "2023-03-02", "data_com": "2023-03-10"},
    {"emissor": "XXXX", "classe": "ON", "tipo": "JRS CAP PROPRIO", "valor": "0.5",
     "aprovado_em": "2023-08-11", "data_com": "2023-08-10"},
    {"emissor": "XXXX", "classe": "ON", "tipo": "JRS CAP PROPRIO", "valor": "0.5",
     "aprovado_em": "2023-08-14", "data_com": "2023-08-10"},
]  # fmt: skip


def test_credito_pelo_historico_sem_data_de_pagamento() -> None:
    """Histórico da aba "Proventos em dinheiro" da B3 (sem pagamento; amostra da Itaúsa): mesma
    empresa e classe, valor dentro de 1% e a data com mais recente antes do pagamento."""
    classes = {"ITSA4": "PN", "ITSA3": "ON", "XXXX3": "ON"}
    linhas = [
        _linha("ITSA4", date(2025, 10, 1), "0.0235295"),  # data com 29/08 (não a de 28/11)
        _linha("ITSA4", date(2026, 3, 2), "0.01820000", 2),  # creditado em 2025, pago em 2026
        _linha("ITSA3", date(2026, 3, 2), "0.0182", 3),  # ON: o histórico só tem PN
        _linha("ITSA4", date(2025, 8, 29), "0.0235295", 4),  # pago na data com: não é dele
        _linha("XXXX3", date(2023, 4, 1), "0.5", 5),  # aprovações antes da mesma data com
        _linha("XXXX3", date(2023, 9, 1), "0.5", 6),  # aprovações depois da data com: empate
    ]
    saida = creditos_b3(linhas, [], {}, None, HISTORICO, classes)
    assert [(x.credito, x.credito_origem) for x in saida] == [
        (date(2025, 8, 29), "B3, histórico"),  # aprovado em 10/02: crédito na data com
        (date(2025, 12, 9), "B3, histórico"),
        (None, ""),
        (None, ""),
        (date(2023, 3, 10), "B3, histórico"),
        (None, ""),
    ]


def test_pacote_com_pagamento_vale_antes_do_historico() -> None:
    [x] = creditos_b3([_linha("ITSA4", date(2025, 3, 10), "0.0242425")], PUBLICOS, {}, None,
                      [dict(HISTORICO[0], valor="0.0242425", data_com="2025-01-02")],
                      {"ITSA4": "PN"})  # fmt: skip
    assert (x.credito, x.credito_origem) == (date(2024, 12, 19), "B3")


BRADESCO = [
    {"emissor": "BBDC", "isin": "BRBBDCACNOR1", "tipo": "JRS CAP PROPRIO", "valor": "0.017249826",
     "aprovado_em": "2024-12-05", "data_com": "2025-12-01", "pagamento": "2026-01-02"},
    {"emissor": "BBDC", "isin": "BRBBDCACNOR1", "tipo": "JRS CAP PROPRIO", "valor": "0.017249826",
     "aprovado_em": "2025-12-18", "data_com": "2026-03-02", "pagamento": "2026-04-01"},
    {"emissor": "BBDC", "isin": "BRBBDCACNOR1", "tipo": "JRS CAP PROPRIO", "valor": "0.351190748",
     "aprovado_em": "2025-12-18", "data_com": "2025-12-29", "pagamento": "2026-07-31"},
]  # fmt: skip


def test_jcp_mensal_aprovado_no_ano_anterior_credita_na_data_com() -> None:
    """Parecer do especialista de 21/10/2026 (dados públicos do Bradesco): o programa mensal
    aprovado em dezembro não é crédito do ano da aprovação. Caso 1b: aprovado em 05/12/2024, data
    com 01/12/2025, pago em 02/01/2026: crédito em 2025 (99.07 em 31/12/2025). Caso 1: aprovado em
    18/12/2025, data com 02/03/2026, pago em 01/04/2026: crédito em 2026, sem 99.07. Caso 2:
    complementar com data com 29/12/2025 e pago em 31/07/2026: crédito em 2025."""
    linhas = [_linha("BBDC3", date(2026, 1, 2), "0.017249826"),
              _linha("BBDC3", date(2026, 4, 1), "0.017249826", 2),
              _linha("BBDC3", date(2026, 7, 31), "0.351190748", 3)]  # fmt: skip
    saida = creditos_b3(linhas, BRADESCO, {})
    assert [x.credito for x in saida] == [date(2025, 12, 1), date(2026, 3, 2), date(2025, 12, 29)]
