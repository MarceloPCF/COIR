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
    itsa = _linha("ITSA4", date(2025, 3, 10), "0.02424250")
    on, pn = _linha("BBDC3", date(2026, 4, 30), "0.3238587", 2), _linha("BBDC4",
                                                                       date(2026, 4, 30),
                                                                       "0.3562445", 3)  # fmt: skip
    outro_dia = _linha("ITSA4", date(2025, 3, 11), "0.0242425", 4)
    informado = _linha("ITSA4", date(2025, 3, 10), "0.0242425", 5).model_copy(
        update={"credito": date(2024, 12, 1), "credito_origem": "informada"}
    )
    saida = creditos_b3([itsa, on, pn, outro_dia, informado], PUBLICOS, {"BBDC3": "BBDC"})
    assert [(x.credito, x.credito_origem) for x in saida] == [
        (date(2024, 12, 16), "B3"),
        (date(2025, 9, 30), "B3"),  # ON: o valor mais próximo
        (date(2025, 11, 30), "B3"),  # PN (emissor pela raiz do ticker)
        (None, ""),  # pagamento em outra data: sem crédito
        (date(2024, 12, 1), "informada"),  # a correção do usuário vale
    ]


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


def test_correcao_do_usuario_muda_o_ano_do_jcp(tmp_path: Path) -> None:
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
