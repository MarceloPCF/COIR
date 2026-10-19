"""F7a: planilhas da Área do Investidor da B3 (geradas no teste, sem dados reais)."""

from collections.abc import Mapping, Sequence
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import Workbook

from coir.domain.b3 import TipoPlanilha
from coir.erros import ImportacaoError
from coir.importers import b3
from coir.servicos.b3 import importar_planilhas
from coir.servicos.cadastros import adicionar_titular
from coir.servicos.estado import Estado
from tests.apoio import cpf_de_teste

D = Decimal
AGORA = datetime(2026, 10, 19, 12, 0, 0)
CAB_MOV = ["Entrada/Saída", "Data", "Movimentação", "Produto", "Instituição", "Quantidade",
           "Preço unitário", "Valor da Operação"]  # fmt: skip
CAB_NEG = ["Data do Negócio", "Tipo de Movimentação", "Mercado", "Prazo/Vencimento",
           "Instituição", "Código de Negociação", "Quantidade", "Preço", "Valor"]  # fmt: skip
CAB_ACOES = ["Produto", "Instituição", "Conta", "Código de Negociação", "CNPJ da Empresa",
             "Código ISIN / Distribuição", "Tipo", "Escriturador", "Quantidade"]  # fmt: skip
CAB_OPCOES = ["Produto", "Instituição", "Tipo de opção", "Variável", "Código de Negociação",
              "Série", "Condição", "Quantidade"]  # fmt: skip
CAB_EMPRESTIMOS = ["Produto", "Instituição", "Natureza", "Número de Contrato", "Quantidade"]
XP = "XP INVESTIMENTOS CCTVM S/A"


def planilha(caminho: Path, abas: Mapping[str, Sequence[Sequence[object]]]) -> Path:
    livro = Workbook()
    livro.remove(livro.active)  # type: ignore[arg-type]
    for nome, linhas in abas.items():
        folha = livro.create_sheet(nome)
        for linha in linhas:
            folha.append(list(linha))
    livro.save(caminho)
    return caminho


def movimentacao(pasta: Path, nome: str = "movimentacao-2026-10-15.xlsx",
                 extra: list[list[object]] | None = None) -> Path:  # fmt: skip
    linhas: list[Sequence[object]] = [
        CAB_MOV,
        ["Credito", "15/09/2026", "Dividendo", "PETR4 - PETROLEO BRASILEIRO S/A PETROBRAS", XP,
         100, 0.5, 50],
        ["Credito", "15/09/2026", "Dividendo", "PETR4 - PETROLEO BRASILEIRO S/A PETROBRAS", XP,
         100, 0.5, 50],  # igual à anterior: legítima, não pode sumir
        ["Credito", "01/09/2026", "Juros Sobre Capital Próprio", "ITSA4 - ITAUSA S/A", XP, "1.000",
         "-", "-"],
        ["Debito", "02/09/2026", "Transferência - Liquidação", "Opção de Compra - PETRK300", XP,
         200, 1.25, 250],
        *(extra or []),
    ]  # fmt: skip
    return planilha(pasta / nome, {"Movimentação": linhas})


def test_movimentacao_lida_com_codigo_numero_e_vazio(tmp_path: Path) -> None:
    linhas = b3.movimentacao(movimentacao(tmp_path))
    assert b3.tipo(tmp_path / "movimentacao-2026-10-15.xlsx") is TipoPlanilha.MOVIMENTACAO
    assert [x.codigo for x in linhas] == ["PETR4", "PETR4", "ITSA4", "PETRK300"]
    assert linhas[2].quantidade == D(1000) and linhas[2].preco is None and linhas[2].valor is None
    assert linhas[0].preco == D("0.5") and linhas[3].entrada is False


def test_negociacao_e_cabecalho_alterado(tmp_path: Path) -> None:
    caminho = planilha(tmp_path / "negociacao.xlsx", {"Negociação": [
        CAB_NEG,
        ["14/08/2026", "Venda", "Mercado à Vista", "-", XP, "CSMG3", 100, 55.47, 5547],
        ["10/03/2026", "Compra", "Futuro", "15/04/2026", XP, "WINJ26", 5, 130000, 650000],
    ]})  # fmt: skip
    [venda, futuro] = b3.negociacao(caminho)
    assert (venda.cv, venda.vencimento, venda.preco) == ("V", None, D("55.47"))
    assert futuro.vencimento == date(2026, 4, 15)
    ruim = planilha(tmp_path / "ruim.xlsx", {"Negociação": [CAB_NEG[:-1]]})
    with pytest.raises(ImportacaoError, match="colunas ausentes Valor"):
        b3.negociacao(ruim)


def test_posicao_do_consolidado_anual(tmp_path: Path) -> None:
    caminho = planilha(tmp_path / "relatorio-consolidado-anual-2025.xlsx", {
        "Posição - Ações": [CAB_ACOES,
                            ["ITSA4 - ITAUSA", XP, "123", "ITSA4", "-", "BRITSAACNPR7", "PN",
                             "ITAU", 1500],
                            [None, None, None, None, None, None, None, None, "Total"]],
        "Posição - Empréstimos": [CAB_EMPRESTIMOS, ["ITSA4 - ITAUSA", XP, "Doador", "1", 500]],
        "Posição - Opções": [CAB_OPCOES, ["Opção", XP, "Compra", "Opção", "ITSAL110", "L110",
                                          "LANÇADOR", 1000]],
    })  # fmt: skip
    posicoes = b3.posicao(caminho)
    assert {p.data for p in posicoes} == {date(2025, 12, 31)}
    assert [(p.origem, p.codigo, p.quantidade) for p in posicoes] == [
        ("acao", "ITSA4", D(1500)), ("emprestimo", "ITSA4", D(500)),
        ("opcao", "ITSAL110", D(-1000))]  # fmt: skip


@pytest.mark.parametrize(
    ("nome", "data"),
    [("relatorio-consolidado-mensal-2026-setembro.xlsx", date(2026, 9, 30)),
     ("relatorio-consolidado-mensal-2024-fevereiro.xlsx", date(2024, 2, 29)),
     ("relatorio-consolidado-mensal-2025-dezembro.xlsx", date(2025, 12, 31))],
)  # fmt: skip
def test_data_da_posicao_pelo_nome(nome: str, data: date) -> None:
    assert b3.data_do_consolidado(Path(nome)) == data


def test_planilha_desconhecida_e_recusada(tmp_path: Path) -> None:
    caminho = planilha(tmp_path / "x.xlsx", {"Outra": [["a"]]})
    with pytest.raises(ImportacaoError, match="não é planilha"):
        b3.tipo(caminho)


def test_importar_nao_duplica_exportacoes_sobrepostas(tmp_path: Path) -> None:
    estado = Estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    antiga = movimentacao(tmp_path, "movimentacao-antiga.xlsx")
    nova = movimentacao(tmp_path, "movimentacao-nova.xlsx", extra=[
        ["Credito", "20/09/2026", "Rendimento", "KNRI11 - KINEA RENDA IMOBILIARIA", XP, 10, 1,
         10]])  # fmt: skip
    rel = importar_planilhas(estado.banco, [antiga], 1, AGORA)
    assert [(a.tipo, a.linhas, a.novas) for a in rel.arquivos] == [
        (TipoPlanilha.MOVIMENTACAO, 4, 4)]  # fmt: skip
    rel = importar_planilhas(estado.banco, [nova, antiga], 1, AGORA)
    assert [(a.linhas, a.novas, a.detalhe) for a in rel.arquivos] == [
        (5, 1, "01/09/2026 a 20/09/2026"), (0, 0, "já importado")]  # fmt: skip


def test_reimportar_a_posicao_da_mesma_data_substitui(tmp_path: Path) -> None:
    estado = Estado(tmp_path / "coir.db")
    estado.preparar(AGORA)
    adicionar_titular(estado.banco, cpf_de_teste(), "Titular")
    nome = "relatorio-consolidado-anual-2025.xlsx"
    linha = ["ITSA4 - ITAUSA", XP, "123", "ITSA4", "-", "BRITSAACNPR7", "PN", "ITAU", 1500]
    (tmp_path / "a").mkdir()
    (tmp_path / "b").mkdir()
    planilha(tmp_path / "a" / nome, {"Posição - Ações": [CAB_ACOES, linha]})
    planilha(tmp_path / "b" / nome, {"Posição - Ações": [CAB_ACOES, [*linha[:-1], 1600]]})
    importar_planilhas(estado.banco, [tmp_path / "a" / nome], 1, AGORA)
    rel = importar_planilhas(estado.banco, [tmp_path / "b" / nome], 1, AGORA)
    assert rel.arquivos[0].detalhe == "posição em 31/12/2025"
    from sqlalchemy import select

    from coir.db.modelos import PosicaoB3
    from coir.db.sessao import conectar

    with conectar(estado.banco) as fabrica, fabrica() as s:
        assert list(s.scalars(select(PosicaoB3.quantidade))) == [D(1600)]
