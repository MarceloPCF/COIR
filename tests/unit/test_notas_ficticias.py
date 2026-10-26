"""Gerador de notas fictícias (tools/notas_ficticias.py): cada nota do roteiro de demonstração sai
em PDF, é lida pelo extrator real e fecha na conciliação; o banco de demonstração apura o ano."""

from datetime import date
from decimal import Decimal as D
from pathlib import Path

import pytest

from coir.domain.fiscal import Categoria
from coir.parsing.extrair import extrair_pdf
from coir.servicos.darf import Situacao
from coir.servicos.estado import Estado
from tests.apoio import cpf_de_teste
from tools import notas_ficticias
from tools.banco_demo import AGORA, criar


@pytest.fixture(scope="module")
def roteiro() -> notas_ficticias.Roteiro:
    return notas_ficticias.carregar()


@pytest.fixture(scope="module")
def estado(tmp_path_factory: pytest.TempPathFactory) -> Estado:
    banco = criar(tmp_path_factory.mktemp("demo"), roteiro=notas_ficticias.ROTEIRO_DEMONSTRACAO)
    estado = Estado(banco)
    estado.preparar(AGORA)
    estado.titular, estado.ano = 1, 2025
    return estado


def test_cada_nota_gerada_e_extraida_e_conciliada(
    roteiro: notas_ficticias.Roteiro, tmp_path: Path
) -> None:
    pdfs, planilha = notas_ficticias.gerar(roteiro, tmp_path)
    assert len(pdfs) == len(roteiro.notas) and planilha is not None
    for caminho, nota in zip(pdfs, roteiro.notas, strict=True):
        res = extrair_pdf(caminho)
        assert not res.rejeitadas, (caminho.name, [r.motivo for r in res.rejeitadas])
        [lida] = res.notas
        assert (lida.numero, lida.data_pregao, lida.conta) == (
            str(nota.numero),
            nota.data,
            "1234567",
        )
        negocios = lida.negocios_bovespa or lida.negocios_bmf
        assert len(negocios) == len(nota.negocios)
        assert [c.cpf for c in res.clientes.values()] == [cpf_de_teste()]


def test_data_sem_pregao_falha(tmp_path: Path) -> None:
    roteiro = tmp_path / "r.yaml"
    texto = notas_ficticias.ROTEIRO_DEMONSTRACAO.read_text(encoding="utf-8")
    roteiro.write_text(texto.replace("data: 2025-01-10", "data: 2025-01-11"), encoding="utf-8")
    with pytest.raises(ValueError, match="11/01/2025 não é dia de pregão"):
        notas_ficticias.carregar(roteiro)


def test_banco_de_demonstracao_apura_o_ano(estado: Estado) -> None:
    """O ano do roteiro: venda isenta em fevereiro, exercício tributável em março, day trade de
    ações (abril, outubro) e de mini-índice (junho, setembro com prejuízo), FII em julho, prejuízo
    de maio compensado em agosto; a DARF de outubro fica em aberto."""
    meses = {m.mes: m for m in estado.apuracao(date(2025, 12, 31)).meses if m.ano == 2025}
    comum = {mes: next(c for c in m.categorias if c.categoria is Categoria.COMUM)
             for mes, m in meses.items()}  # fmt: skip
    assert comum[2].isento > 0 and comum[2].imposto == 0
    assert comum[5].prejuizo_a_compensar < 0
    assert comum[8].prejuizo_anterior < 0 < comum[8].imposto
    darfs = {d.mes: d for d in estado.darfs_todos_os_anos() if d.ano == 2025}
    assert sorted(darfs) == [3, 4, 6, 7, 8, 10]
    assert {m for m, d in darfs.items() if d.situacao is Situacao.PAGA} == {3, 4, 6, 7, 8}
    assert darfs[10].situacao is Situacao.EM_ATRASO and darfs[10].principal > D(10)


def test_opcoes_e_proventos_da_demonstracao(estado: Estado) -> None:
    ciclos = {c.ticker: c.situacao for c in estado.opcoes()}
    assert ciclos == {"PETRC380": "exercida", "VALEQ550": "vencida (pó)"}
    proventos, avisos = estado.proventos()
    assert len([p for p in proventos if p.data.year == 2025]) == 9 and not avisos
