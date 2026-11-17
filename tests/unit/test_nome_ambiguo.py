"""Nome de pregão que valia para mais de um código na data (dois BDR da mesma companhia, por
exemplo): a nota não diz qual. O COIR desempata pela planilha de negociação da B3 do titular
(mesmo dia, C/V, quantidade e preço); sem ela, pergunta na tela Importar e guarda a escolha para
a conta. Nunca escolhe no palpite (decisão de 16/11/2026)."""

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest
from sqlalchemy import delete, select

from coir import referencias
from coir.db.modelos import CodigoEscolhido, NegociacaoB3, Operacao
from coir.db.sessao import conectar, migrar
from coir.parsing.extrair import ResultadoExtracao, extrair_pdf
from coir.servicos.b3 import importar_planilhas
from coir.servicos.cadastros import adicionar_titular
from coir.servicos.importacao import (
    NomeAmbiguo,
    RelatorioImportacao,
    escolher_codigo,
    importar_lista,
)
from tests.apoio import cpf_de_teste
from tests.unit.test_importador_b3 import CAB_NEG, XP, planilha
from tools import notas_ficticias

AGORA = datetime(2026, 11, 16, 12, 0)
NOME = "EMPRESA Y DR1"
REFS = referencias.carregar().model_copy(update={"nomes_historicos": {
    "empresaydr1": (("2021-10-04", "2026-11-09", "EMPY31"), ("2023-07-03", "2025-09-12", "EMPY20")),
}})  # fmt: skip


@pytest.fixture
def pdf(tmp_path: Path) -> Path:
    nome, texto = notas_ficticias.textos(notas_ficticias.carregar())[0]
    return notas_ficticias.pdf(texto, tmp_path / f"{nome}.pdf")


def _com_nome_em_duvida(pdf: Path) -> ResultadoExtracao:
    """A nota fictícia (10/01/2025) com o terceiro negócio, 200 a 52,10, no nome em dúvida."""
    res = extrair_pdf(pdf)
    [nota] = res.notas
    negocios = list(nota.negocios_bovespa)
    negocios[2] = negocios[2].model_copy(update={"especificacao": NOME})
    res.notas = [nota.model_copy(update={"negocios_bovespa": tuple(negocios)})]
    return res


def _banco(tmp_path: Path, pdf: Path) -> tuple[Path, int]:
    banco = tmp_path / "coir.db"
    migrar(banco, AGORA)
    cpf = next(iter(extrair_pdf(pdf).clientes.values())).cpf
    return banco, adicionar_titular(banco, cpf, "Titular")


def _importar(banco: Path, pdf: Path) -> RelatorioImportacao:
    return importar_lista(banco, [pdf], REFS, AGORA, extrator=_com_nome_em_duvida)


def _negociacao(tmp_path: Path, banco: Path, titular: int, *linhas: tuple[str, int, float]) -> None:
    corpo = [["10/01/2025", "Compra", "Mercado à Vista", "-", XP, codigo, qtd, preco, qtd * preco]
             for codigo, qtd, preco in linhas]  # fmt: skip
    arquivo = planilha(tmp_path / f"negociacao-{len(linhas)}-{linhas[0][0]}.xlsx",
                       {"Negociação": [CAB_NEG, *corpo]})  # fmt: skip
    importar_planilhas(banco, [arquivo], titular, AGORA)


def _papeis(banco: Path) -> list[str]:
    with conectar(banco) as fabrica, fabrica() as s:
        return list(s.scalars(select(Operacao.papel).order_by(Operacao.ordem)))


def _escolher(banco: Path, duvida: NomeAmbiguo, ticker: str, conta: str | None = None) -> None:
    escolher_codigo(banco, duvida.corretora, conta or duvida.conta, duvida.nome, duvida.data,
                    ticker, AGORA)  # fmt: skip


def test_sem_planilha_nem_escolha_a_nota_e_recusada_e_a_duvida_vai_para_a_tela(
    tmp_path: Path, pdf: Path
) -> None:
    banco, _ = _banco(tmp_path, pdf)
    rel = _importar(banco, pdf)
    assert rel.notas_gravadas == 0 and len(rel.rejeitadas) == 1
    assert "mais de um código (EMPY20, EMPY31)" in rel.rejeitadas[0].motivo
    [duvida] = rel.ambiguos
    assert (duvida.especificacao, duvida.nome, duvida.codigos, duvida.data) == (
        NOME, "empresaydr1", ("EMPY20", "EMPY31"), date(2025, 1, 10))  # fmt: skip
    assert _papeis(banco) == []


def test_a_planilha_de_negociacao_da_b3_aponta_o_codigo(tmp_path: Path, pdf: Path) -> None:
    """Mesmo titular, dia, C/V, quantidade e preço em um só dos códigos: é ele. O fracionário
    (final F) da planilha é o mesmo ativo."""
    banco, titular = _banco(tmp_path, pdf)
    _negociacao(tmp_path, banco, titular, ("EMPY20F", 200, 52.1), ("EMPY31", 300, 48.0))
    rel = _importar(banco, pdf)
    assert rel.notas_gravadas == 1 and not rel.rejeitadas and not rel.ambiguos
    assert _papeis(banco)[2] == "EMPY20"


def test_planilha_que_reparte_o_negocio_vale_pela_soma_do_dia(tmp_path: Path, pdf: Path) -> None:
    banco, titular = _banco(tmp_path, pdf)
    _negociacao(tmp_path, banco, titular, ("EMPY31", 150, 52.1), ("EMPY31", 50, 52.1))
    assert _importar(banco, pdf).notas_gravadas == 1
    assert _papeis(banco)[2] == "EMPY31"


@pytest.mark.parametrize(
    "linhas",
    [
        (("EMPY20", 200, 52.1), ("EMPY31", 200, 52.1)),  # os dois códigos batem
        (("EMPY20", 200, 52.2),),  # outro preço
        (("EMPY20", 100, 52.1),),  # outra quantidade
        (("OUTR3", 200, 52.1),),  # código que não é candidato
    ],
)
def test_planilha_que_nao_aponta_um_so_codigo_nao_decide(
    tmp_path: Path, pdf: Path, linhas: tuple[tuple[str, int, float], ...]
) -> None:
    banco, titular = _banco(tmp_path, pdf)
    _negociacao(tmp_path, banco, titular, *linhas)
    rel = _importar(banco, pdf)
    assert rel.notas_gravadas == 0 and len(rel.ambiguos) == 1


def test_planilha_de_outro_titular_nao_decide(tmp_path: Path, pdf: Path) -> None:
    banco, _ = _banco(tmp_path, pdf)
    outro = adicionar_titular(banco, cpf_de_teste("111444777"), "Outro")
    _negociacao(tmp_path, banco, outro, ("EMPY20", 200, 52.1))
    assert len(_importar(banco, pdf).ambiguos) == 1


def test_a_escolha_do_usuario_fica_guardada_para_a_conta(tmp_path: Path, pdf: Path) -> None:
    banco, _ = _banco(tmp_path, pdf)
    [duvida] = _importar(banco, pdf).ambiguos
    _escolher(banco, duvida, "EMPY31")
    rel = _importar(banco, pdf)  # o arquivo ficou incompleto: é lido de novo
    assert rel.notas_gravadas == 1 and not rel.ambiguos
    assert _papeis(banco)[2] == "EMPY31"
    _escolher(banco, duvida, "EMPY20")
    with conectar(banco) as fabrica, fabrica() as s:
        assert list(s.scalars(select(CodigoEscolhido.ticker))) == ["EMPY20"]  # troca, não soma


def test_escolha_de_outra_conta_ou_fora_dos_candidatos_nao_vale(tmp_path: Path, pdf: Path) -> None:
    banco, _ = _banco(tmp_path, pdf)
    [duvida] = _importar(banco, pdf).ambiguos
    _escolher(banco, duvida, "EMPY31", "outra-conta")
    _escolher(banco, duvida, "OUTR3")
    assert len(_importar(banco, pdf).ambiguos) == 1


def test_escolha_de_outro_pregao_nao_vale(tmp_path: Path, pdf: Path) -> None:
    """Os dois códigos podem ter sido negociados em dias diferentes: a escolha vale só para o
    pregão em que foi feita."""
    banco, _ = _banco(tmp_path, pdf)
    [duvida] = _importar(banco, pdf).ambiguos
    escolher_codigo(banco, duvida.corretora, duvida.conta, duvida.nome, date(2025, 1, 9),
                    "EMPY31", AGORA)  # fmt: skip
    assert len(_importar(banco, pdf).ambiguos) == 1


def test_escolha_nova_nao_troca_o_codigo_da_nota_ja_gravada(tmp_path: Path, pdf: Path) -> None:
    """Revisão de 16/11/2026: a nota entrou com EMPY31; trocar a escolha para EMPY20 não regrava
    as operações em silêncio (o preço médio mudaria sem ninguém pedir)."""
    banco, _ = _banco(tmp_path, pdf)
    [duvida] = _importar(banco, pdf).ambiguos
    escolher_codigo(banco, duvida.corretora, duvida.conta, duvida.nome, duvida.data, "EMPY31",
                    AGORA)  # fmt: skip
    _importar(banco, pdf)
    escolher_codigo(banco, duvida.corretora, duvida.conta, duvida.nome, duvida.data, "EMPY20",
                    AGORA)  # fmt: skip
    rel = _importar(banco, pdf)
    assert rel.notas_renormalizadas == 0
    assert _papeis(banco)[2] == "EMPY31"


def test_a_planilha_da_b3_vale_sobre_a_escolha_e_corrige_a_nota_gravada(
    tmp_path: Path, pdf: Path
) -> None:
    """A escolha foi EMPY31; a planilha importada depois mostra EMPY20: as operações da nota já
    gravada são regravadas na importação seguinte (o livro, a nota, não muda)."""
    banco, titular = _banco(tmp_path, pdf)
    [duvida] = _importar(banco, pdf).ambiguos
    _escolher(banco, duvida, "EMPY31")
    _importar(banco, pdf)
    _negociacao(tmp_path, banco, titular, ("EMPY20", 200, 52.1))
    rel = _importar(banco, pdf)
    assert rel.notas_renormalizadas == 1
    assert _papeis(banco)[2] == "EMPY20"


def test_nota_gravada_continua_com_o_codigo_se_a_planilha_some(tmp_path: Path, pdf: Path) -> None:
    """A nota entrou pelo código da planilha; sem as linhas dela, a conferência das notas gravadas
    não pode travar as importações seguintes nem trocar o código."""
    banco, titular = _banco(tmp_path, pdf)
    _negociacao(tmp_path, banco, titular, ("EMPY20", 200, 52.1))
    _importar(banco, pdf)
    with conectar(banco) as fabrica, fabrica.begin() as s:
        s.execute(delete(NegociacaoB3))
    rel = _importar(banco, pdf)
    assert rel.notas_renormalizadas == 0 and not rel.rejeitadas
    assert _papeis(banco)[2] == "EMPY20"


def test_preco_da_planilha_comparado_como_numero(tmp_path: Path, pdf: Path) -> None:
    """52,1 na planilha e 52,10 na nota são o mesmo preço."""
    banco, titular = _banco(tmp_path, pdf)
    _negociacao(tmp_path, banco, titular, ("EMPY20", 200, 52.1))
    with conectar(banco) as fabrica, fabrica() as s:
        assert s.scalar(select(NegociacaoB3.preco)) == Decimal("52.10")
    assert _importar(banco, pdf).notas_gravadas == 1
