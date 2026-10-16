"""F3: todas as notas reais (COIR_DADOS) importadas no banco; reimportar não muda nada; o banco
reproduz exatamente a extração (mesmo candidato do gabarito). Critério de saída da F3 (PLANO §8).

O banco é temporário (tmp_path), nunca o coir.db do usuário.
"""

import time
from datetime import datetime
from pathlib import Path

import pytest
from sqlalchemy.orm import Session, sessionmaker

from coir import referencias
from coir.config import dados_dir
from coir.db import repositorio as repo
from coir.db.sessao import abrir, migrar
from coir.domain.nota import NotaExtraida, OperacaoExtraida
from coir.parsing.extrair import extrair_pdf
from coir.parsing.normalizar import classes_das_series, operacoes
from coir.servicos.importacao import RelatorioImportacao, arquivos_da_pasta, importar_arquivos
from tests.apoio import cpf_de_teste

pytestmark = pytest.mark.golden
NOTAS = dados_dir() / "_notas-corretagens"
AGORA = datetime(2026, 10, 14, 12, 0, 0)


@pytest.fixture(scope="module")
def fabrica(tmp_path_factory: pytest.TempPathFactory) -> sessionmaker[Session]:
    if not NOTAS.is_dir():
        pytest.skip("pasta de notas ausente")
    banco = tmp_path_factory.mktemp("banco") / "coir.db"
    migrar(banco, AGORA)
    f = abrir(banco)
    with f.begin() as s:
        repo.adicionar_titular(s, cpf_de_teste(), "Titular de Teste")
    return f


@pytest.fixture(scope="module")
def arquivos() -> list[Path]:
    return arquivos_da_pasta(NOTAS, RelatorioImportacao())


@pytest.fixture(scope="module")
def primeira(fabrica: sessionmaker[Session], arquivos: list[Path]) -> RelatorioImportacao:
    with fabrica() as s:
        return importar_arquivos(s, arquivos, referencias.carregar(), AGORA)


def test_todas_as_notas_importadas(primeira: RelatorioImportacao) -> None:
    assert primeira.rejeitadas == []
    assert primeira.notas_gravadas == 2034


def test_reimportar_nao_muda_nada_e_e_rapido(
    fabrica: sessionmaker[Session], arquivos: list[Path], primeira: RelatorioImportacao
) -> None:
    with fabrica() as s:
        antes = repo.contagens(s)
    inicio = time.perf_counter()
    with fabrica() as s:
        rel = importar_arquivos(s, arquivos, referencias.carregar(), AGORA)
    segundos = time.perf_counter() - inicio
    assert not rel.mudou and rel.arquivos_lidos == [] and rel.rejeitadas == []
    with fabrica() as s:
        assert repo.contagens(s) == antes
    assert segundos < 5, f"reimportação levou {segundos:.1f} s"


def test_banco_reproduz_a_extracao(
    fabrica: sessionmaker[Session], arquivos: list[Path], primeira: RelatorioImportacao
) -> None:
    refs = referencias.carregar()
    resultados = [extrair_pdf(a) for a in arquivos]
    notas = [n for r in resultados for n in r.notas]
    clientes = {k for r in resultados for k in r.clientes}
    assert [n.arquivo for n in notas if n.chave not in clientes] == []  # CPF impresso em todas
    series = classes_das_series(notas, refs)
    esperado: dict[tuple[object, ...], tuple[NotaExtraida, list[OperacaoExtraida]]] = {}
    for n in notas:  # nota em dois PDFs: fica a do primeiro arquivo, como na importação
        esperado.setdefault(n.chave, (n, operacoes(n, refs, series)))
    with fabrica() as s:
        lidas: dict[tuple[object, ...], tuple[NotaExtraida, list[OperacaoExtraida]]] = {
            n.chave: (n, ops) for n, ops in repo.carregar_notas(s)
        }
    assert lidas.keys() == esperado.keys()
    assert [k for k in esperado if lidas[k] != esperado[k]] == []


def test_gravar_500_notas_ja_extraidas_em_segundos(tmp_path: Path, arquivos: list[Path]) -> None:
    refs = referencias.carregar()
    escolhidos: list[Path] = []
    extraidas: dict[str, object] = {}
    for a in arquivos:
        res = extrair_pdf(a)
        extraidas[a.name] = res
        escolhidos.append(a)
        if sum(len(r.notas) for r in extraidas.values()) >= 500:  # type: ignore[attr-defined]
            break
    banco = tmp_path / "coir.db"
    migrar(banco, AGORA)
    f = abrir(banco)
    with f.begin() as s:
        repo.adicionar_titular(s, cpf_de_teste(), "Titular de Teste")
    inicio = time.perf_counter()
    with f() as s:
        rel = importar_arquivos(s, escolhidos, refs, AGORA, extrator=lambda p: extraidas[p.name])  # type: ignore[arg-type, return-value]
    segundos = time.perf_counter() - inicio
    assert rel.notas_gravadas >= 500 and rel.rejeitadas == []
    assert segundos < 5, f"gravar {rel.notas_gravadas} notas levou {segundos:.1f} s"
