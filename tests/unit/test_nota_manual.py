"""Nota de corretagem digitada à mão (pedido do usuário de 24/11/2026): a nota perdida ou de
corretora sem nota SINACOR entra no livro como as lidas do PDF e passa pelas mesmas regras. Nenhuma
regra fiscal nova: a prova é que a nota digitada dá a mesma apuração que a mesma nota lida do PDF
(roteiro de demonstração, com notas Bovespa e BM&F, opções, day trade e ajuste de posição)."""

from dataclasses import replace
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest
from sqlalchemy import func, select

from coir import referencias
from coir.db.modelos import ArquivoImportado, DarfPagamento, NotaAnterior, NotaCorretagem, Titular
from coir.db.sessao import conectar, migrar
from coir.domain.nota import CategoriaCusto, NotaExtraida, Segmento
from coir.erros import ImportacaoError
from coir.parsing.extrair import extrair_pdf
from coir.servicos import nota_manual as nm
from coir.servicos.apuracao import apurar_banco
from coir.servicos.cadastros import adicionar_titular
from coir.servicos.importacao import importar_lista
from coir.servicos.notas import detalhe
from tests.apoio import cpf_de_teste
from tools import notas_ficticias

AGORA = datetime(2026, 11, 24, 12, 0)
REFS = referencias.carregar()
FISCAIS = referencias.carregar_fiscais()


@pytest.fixture(scope="module")
def pdfs(tmp_path_factory: pytest.TempPathFactory) -> list[Path]:
    pasta = tmp_path_factory.mktemp("notas")
    return [notas_ficticias.pdf(texto, pasta / f"{nome}.pdf")
            for nome, texto in notas_ficticias.textos(notas_ficticias.carregar())]  # fmt: skip


@pytest.fixture(scope="module")
def extraidas(pdfs: list[Path]) -> list[NotaExtraida]:
    return [n for p in pdfs for n in extrair_pdf(p).notas]


def _titular(banco: Path) -> int:
    migrar(banco, AGORA)
    return adicionar_titular(banco, cpf_de_teste(), "Titular de Teste")


def _resumo(banco: Path) -> list[tuple[object, ...]]:
    apuracao = apurar_banco(banco, REFS, FISCAIS, date(2026, 11, 1))
    return [(m.ano, m.mes, m.imposto, m.darf, m.irrf_mes,
             tuple((c.categoria, c.resultado) for c in m.categorias))
            for m in apuracao.meses]  # fmt: skip


def _gravar_todas(banco: Path, titular: int, notas: list[NotaExtraida]) -> list[int]:
    ids: list[int] = []
    for n in sorted(notas, key=lambda x: x.data_pregao):
        gravada = nm.gravar(banco, titular, nm.de_extraida(n, None), REFS, FISCAIS, AGORA)
        assert gravada.avisos == []
        ids.append(gravada.nota_id)
    return ids


def _contar(banco: Path, modelo: type) -> int:
    with conectar(banco) as fabrica, fabrica() as s:
        return s.scalar(select(func.count()).select_from(modelo)) or 0


def test_nota_digitada_da_a_mesma_apuracao_que_a_do_pdf(
    tmp_path: Path, pdfs: list[Path], extraidas: list[NotaExtraida]
) -> None:
    """As 15 notas do roteiro de demonstração (Bovespa e BM&F), lidas do PDF num banco e digitadas
    noutro: mesmo imposto, DARF, IRRF e resultado por categoria em todos os meses."""
    assert {n.segmento for n in extraidas} == {Segmento.BOVESPA, Segmento.BMF}
    do_pdf = tmp_path / "pdf.db"
    importar_lista(do_pdf, pdfs, REFS, AGORA)
    manual = tmp_path / "manual.db"
    _gravar_todas(manual, _titular(manual), extraidas)
    assert _resumo(manual) == _resumo(do_pdf)
    assert any(m[2] for m in _resumo(manual))  # o roteiro tem imposto: a comparação não é vazia


def test_nota_manual_concilia_e_fica_marcada(tmp_path: Path, extraidas: list[NotaExtraida]) -> None:
    banco = tmp_path / "coir.db"
    [nota_id] = _gravar_todas(banco, _titular(banco), extraidas[:1])
    d = detalhe(banco, tmp_path, nota_id)
    assert nm.eh_manual(d.resumo.arquivo) and d.pdf is None
    assert d.resumo.paginas == ()


def test_corrigir_e_remover_guardam_a_versao_anterior(
    tmp_path: Path, extraidas: list[NotaExtraida]
) -> None:
    banco = tmp_path / "coir.db"
    titular = _titular(banco)
    [nota_id] = _gravar_todas(banco, titular, extraidas[:1])
    formulario = nm.para_corrigir(banco, titular, nota_id)
    linha = formulario.bovespa[0]
    corrigida = replace(formulario, bovespa=(replace(linha, preco=linha.preco + 1),
                                             *formulario.bovespa[1:]))  # fmt: skip
    novo = nm.gravar(banco, titular, corrigida, REFS, FISCAIS, AGORA, substituir=nota_id).nota_id
    assert _contar(banco, NotaCorretagem) == 1 and _contar(banco, NotaAnterior) == 1
    assert nm.para_corrigir(banco, titular, novo).bovespa[0].preco == linha.preco + 1
    assert nm.remover(banco, titular, novo, REFS, FISCAIS, AGORA) == []
    assert _contar(banco, NotaCorretagem) == 0 and _contar(banco, NotaAnterior) == 2
    assert len(list((tmp_path / "backups").glob("*.db"))) >= 2  # backup antes de cada mudança


def test_validacoes(tmp_path: Path, extraidas: list[NotaExtraida]) -> None:
    banco = tmp_path / "coir.db"
    titular = _titular(banco)
    base = nm.de_extraida(extraidas[0], None)
    linha = base.bovespa[0]
    casos = {
        "número": replace(base, numero=" "),
        "maior que zero": replace(base, bovespa=(replace(linha, quantidade=Decimal(0)),)),
        "custo negativo": replace(base, custos={CategoriaCusto.CORRETAGEM: Decimal(-1)}),
        "não tem": replace(base, custos={CategoriaCusto.TAXAS_BMF: Decimal(1)}),
        "MM/AA": replace(base, bovespa=(replace(linha, mercado="OPCAO DE COMPRA", prazo=""),)),
        "corretora e o número": replace(base, corretora=""),
    }
    for trecho, nota in casos.items():
        with pytest.raises(ImportacaoError, match=trecho):
            nm.gravar(banco, titular, nota, REFS, FISCAIS, AGORA)
    assert _contar(banco, NotaCorretagem) == 0
    assert not (tmp_path / "backups").exists() or not list((tmp_path / "backups").glob("*.db"))


def test_codigo_que_as_tabelas_nao_conhecem_entra_com_aviso(
    tmp_path: Path, extraidas: list[NotaExtraida]
) -> None:
    """Final de ação basta para o tipo (como no código impresso pela corretora), mas o código
    digitado pode estar errado: grava e avisa, porque ativo novo pode faltar nas tabelas."""
    banco = tmp_path / "coir.db"
    base = nm.de_extraida(extraidas[0], None)
    nota = replace(base, bovespa=(replace(base.bovespa[0], ativo="ZZZZ3"),))
    gravada = nm.gravar(banco, _titular(banco), nota, REFS, FISCAIS, AGORA)
    assert [a.split(":")[0] for a in gravada.avisos] == ["ZZZZ3"]


def test_mesmo_numero_na_mesma_conta_e_recusado(
    tmp_path: Path, extraidas: list[NotaExtraida]
) -> None:
    banco = tmp_path / "coir.db"
    titular = _titular(banco)
    _gravar_todas(banco, titular, extraidas[:1])
    with pytest.raises(ImportacaoError, match="já existe a nota"):
        _gravar_todas(banco, titular, extraidas[:1])


def test_numero_sugerido(tmp_path: Path, extraidas: list[NotaExtraida]) -> None:
    banco = tmp_path / "coir.db"
    titular = _titular(banco)
    assert nm.proximo_numero(banco, titular) == "M1"
    nota = replace(nm.de_extraida(extraidas[0], None), numero="M1")
    nm.gravar(banco, titular, nota, REFS, FISCAIS, AGORA)
    assert nm.proximo_numero(banco, titular) == "M2"


def test_pdf_do_mesmo_dia_avisa_e_vira_pendencia(
    tmp_path: Path, pdfs: list[Path], extraidas: list[NotaExtraida]
) -> None:
    """A nota do PDF que chega depois da manual não a apaga (regra 7): aviso na importação e
    pendência até o usuário remover a manual."""
    banco = tmp_path / "coir.db"
    primeira, segunda = extraidas[0], extraidas[1]
    pdf_de = {n.numero: next(p for p in pdfs if n.numero in p.name) for n in (primeira, segunda)}
    importar_lista(banco, [pdf_de[segunda.numero]], REFS, AGORA)  # cria o titular do CPF da nota
    with conectar(banco) as fabrica, fabrica() as s:
        titular = s.scalar(select(Titular.id))
    assert titular is not None
    nota = replace(nm.de_extraida(primeira, None), numero="M1")
    nm.gravar(banco, titular, nota, REFS, FISCAIS, AGORA)
    assert nm.sobrepostas(banco, titular) == []
    rel = importar_lista(banco, [pdf_de[primeira.numero]], REFS, AGORA)
    assert rel.rejeitadas == []
    assert any("há nota manual (M1)" in a for a in rel.avisos)
    [sobreposta] = nm.sobrepostas(banco, titular)
    assert (sobreposta.numero, sobreposta.data) == ("M1", primeira.data_pregao)


def test_corrigir_trocando_o_numero_mantem_o_historico_e_avisa_darf_paga(
    tmp_path: Path, extraidas: list[NotaExtraida]
) -> None:
    """A nota corrigida segue no mesmo registro de arquivo (o histórico a acha) e a gravação avisa
    a DARF paga do mês da nota em diante, como a remoção."""
    banco = tmp_path / "coir.db"
    titular = _titular(banco)
    primeira = extraidas[0]
    [nota_id] = _gravar_todas(banco, titular, [primeira])
    with conectar(banco) as fabrica, fabrica.begin() as s:
        s.add(DarfPagamento(
            titular_id=titular, ano=primeira.data_pregao.year, mes=primeira.data_pregao.month,
            data_pagamento=AGORA.date(), valor_pago=Decimal(10), observacao="",
            registrado_em=AGORA,
        ))  # fmt: skip
    nova = replace(nm.para_corrigir(banco, titular, nota_id), numero="M9")
    gravada = nm.gravar(banco, titular, nova, REFS, FISCAIS, AGORA, substituir=nota_id)
    assert any("DARF(s) paga(s)" in a for a in gravada.avisos)
    d = detalhe(banco, tmp_path, gravada.nota_id)
    assert d.resumo.numero == "M9" and nm.eh_manual(d.resumo.arquivo) and "M9" in d.resumo.arquivo
    assert [v.nota.numero for v in d.anteriores] == [primeira.numero]
    assert _contar(banco, ArquivoImportado) == 1  # sem registro órfão


def test_nota_do_pdf_nao_se_corrige_nem_se_remove_aqui(
    tmp_path: Path, pdfs: list[Path], extraidas: list[NotaExtraida]
) -> None:
    banco = tmp_path / "coir.db"
    importar_lista(banco, pdfs[:1], REFS, AGORA)
    with conectar(banco) as fabrica, fabrica() as s:
        titular, nota_id = s.execute(select(Titular.id, NotaCorretagem.id)).one()
    with pytest.raises(ImportacaoError, match="só a nota manual"):
        nm.remover(banco, titular, nota_id, REFS, FISCAIS, AGORA)
    with pytest.raises(ImportacaoError, match="só a nota manual"):
        nm.gravar(banco, titular, nm.de_extraida(extraidas[0], None), REFS, FISCAIS, AGORA,
                  substituir=nota_id)  # fmt: skip
    assert _contar(banco, NotaCorretagem) == 1


def test_nota_que_o_motor_recusa_nao_grava_nada(tmp_path: Path) -> None:
    """Exercício de opção sem a opção na carteira: o motor para, a transação é desfeita e o
    banco fica como estava (só o backup foi feito)."""
    banco = tmp_path / "coir.db"
    titular = _titular(banco)
    nota = nm.NotaManual(
        conta_id=None, corretora="XP", conta="123", data=date(2025, 3, 10), numero="M1",
        segmento=Segmento.BOVESPA,
        bovespa=(nm.LinhaBovespa(cv="C", mercado="EXERC OPC COMPRA", ativo="PETRC300E ON",
                                 quantidade=Decimal(100), preco=Decimal(30)),),
    )  # fmt: skip
    with pytest.raises(ImportacaoError, match="não fecha com as operações"):
        nm.gravar(banco, titular, nota, REFS, FISCAIS, AGORA)
    assert _contar(banco, NotaCorretagem) == 0 and _contar(banco, ArquivoImportado) == 0
