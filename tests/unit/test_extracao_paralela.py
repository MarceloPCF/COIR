"""Leitura do lote em vários processos (importação): o mesmo resultado da leitura um a um, na
ordem dos arquivos (regra 6), inclusive o que a auditoria registra de cada rejeição, e a mesma
importação gravada numa transação só (regra 8)."""

import os
import shutil
from datetime import datetime
from pathlib import Path

import pytest
from sqlalchemy import select

from coir import referencias
from coir.db import repositorio as repo
from coir.db.modelos import NotaCorretagem
from coir.db.sessao import abrir, migrar
from coir.erros import ImportacaoError
from coir.parsing.extrair import ResultadoExtracao, extrair_pdf
from coir.servicos import extracao, trilha
from coir.servicos.importacao import importar_arquivos
from tests.apoio import cpf_de_teste

FIXTURES = Path(__file__).parents[1] / "fixtures" / "notas"
AGORA = datetime(2026, 10, 14, 12, 0, 0)


def _lote(destino: Path) -> list[Path]:
    """As notas das fixtures e um arquivo que não é nota (rejeitado, com erro para a auditoria):
    acima do mínimo para ler em processos."""
    destino.mkdir()
    arquivos = [Path(shutil.copy(p, destino / p.name)) for p in sorted(FIXTURES.glob("*.txt"))]
    ruim = destino / "zz_nao_e_nota.txt"
    ruim.write_text("Relatório qualquer\nsem nota de corretagem\n", encoding="utf-8")
    assert len(arquivos) + 1 >= extracao.MINIMO_PARALELO
    return [*arquivos, ruim]


@pytest.fixture
def dois_processos(monkeypatch: pytest.MonkeyPatch) -> None:
    """Dois processos mesmo numa máquina de dois núcleos (o CI)."""
    monkeypatch.setattr(extracao, "processos", lambda n: 2)


def test_processos_so_com_lote_grande() -> None:
    assert extracao.processos(extracao.MINIMO_PARALELO - 1) == 1
    assert extracao.processos(extracao.MINIMO_PARALELO) >= 1


@pytest.mark.usefixtures("dois_processos")
def test_lote_em_processos_igual_a_leitura_um_a_um(tmp_path: Path) -> None:
    arquivos = _lote(tmp_path / "notas")
    um_a_um = [extrair_pdf(a) for a in arquivos]
    em_processos = list(extracao.extrair_lote(arquivos, ()))
    assert [r.notas for r in em_processos] == [r.notas for r in um_a_um]
    assert [r.avisos for r in em_processos] == [r.avisos for r in um_a_um]
    assert [r.clientes for r in em_processos] == [r.clientes for r in um_a_um]
    rejeitadas = [x for r in um_a_um for x in r.rejeitadas]
    assert rejeitadas, "o lote precisa de uma rejeição com erro"
    assert [x for r in em_processos for x in r.rejeitadas] == rejeitadas
    assert [x.auditoria for r in em_processos for x in r.rejeitadas] == [
        trilha.dados_rejeicao(x) for x in rejeitadas
    ]


def _importar(banco: Path, arquivos: list[Path]) -> tuple[object, ...]:
    migrar(banco, AGORA)
    fabrica = abrir(banco)
    with fabrica.begin() as s:
        repo.adicionar_titular(s, cpf_de_teste(), "Titular de Teste")
    with fabrica() as s:
        rel = importar_arquivos(s, arquivos, referencias.carregar(), AGORA)
    with fabrica() as s:
        consulta = select(NotaCorretagem).order_by(NotaCorretagem.id)
        notas = [(n.numero, n.data_pregao, n.segmento, n.json_nota) for n in s.scalars(consulta)]
        return (rel.notas_gravadas, rel.operacoes_gravadas, len(rel.rejeitadas),
                rel.arquivos_lidos, repo.contagens(s)), notas  # fmt: skip


def test_importacao_em_processos_grava_o_mesmo_banco(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    arquivos = _lote(tmp_path / "notas")
    monkeypatch.setattr(extracao, "processos", lambda n: 1)
    um_a_um = _importar(tmp_path / "um.db", arquivos)
    monkeypatch.setattr(extracao, "processos", lambda n: 2)
    em_processos = _importar(tmp_path / "varios.db", arquivos)
    assert em_processos == um_a_um


def _derruba_no_ultimo(caminho: Path, senhas: tuple[str, ...]) -> ResultadoExtracao:
    """Leitura que derruba o processo (como uma queda nativa do PDFium) no arquivo ruim."""
    if caminho.name.startswith("zz_"):
        os._exit(3)
    return extracao._extrair(caminho, senhas)  # pyright: ignore[reportPrivateUsage]


@pytest.mark.usefixtures("dois_processos")
def test_processo_de_leitura_que_cai_vira_erro_de_importacao(tmp_path: Path) -> None:
    """Um processo de leitura que morre não derruba o programa com BrokenProcessPool: vira erro
    de importação (o lote é desfeito e a tela mostra a mensagem)."""
    arquivos = _lote(tmp_path / "notas")
    with pytest.raises(ImportacaoError, match="processo de leitura"):
        list(extracao.extrair_lote(arquivos, (), ler=_derruba_no_ultimo))


def test_progresso_conta_so_os_arquivos_lidos(tmp_path: Path) -> None:
    """Na reimportação, os arquivos já importados não são lidos: o andamento conta os que serão
    lidos, do primeiro ao último, sem pular."""
    arquivos = _lote(tmp_path / "notas")[:3]
    banco = tmp_path / "coir.db"
    migrar(banco, AGORA)
    fabrica = abrir(banco)
    with fabrica.begin() as s:
        repo.adicionar_titular(s, cpf_de_teste(), "Titular de Teste")
    with fabrica() as s:
        importar_arquivos(s, arquivos[:1], referencias.carregar(), AGORA)
    passos: list[tuple[int, int, str]] = []
    with fabrica() as s:
        importar_arquivos(s, arquivos, referencias.carregar(), AGORA,
                          progresso=lambda *p: passos.append(p))  # fmt: skip
    assert passos == [(0, 2, arquivos[1].name), (1, 2, arquivos[2].name), (2, 2, "")]
