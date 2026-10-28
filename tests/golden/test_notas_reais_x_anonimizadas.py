"""As notas anonimizadas (COIR_DADOS/_notas-corretagens) só valem como base de teste se a extração
delas for a mesma das originais. Na 3.0.1 não era: a anonimização levou a conta da BTG para a
última linha, e o leitor, ajustado às cópias, lia nas originais o telefone da ouvidoria.

Roda com COIR_NOTAS_REAIS apontando para a pasta das notas originais (fora do repositório; só no
PC do mantenedor). Compara nota a nota: as mesmas notas, nada rejeitado e conteúdo igual fora a
conta; cada conta real corresponde a uma só conta anonimizada."""

import os
from pathlib import Path

import pytest

from coir.config import dados_dir
from coir.domain.nota import NotaExtraida
from coir.parsing.extrair import extrair_pasta

pytestmark = [pytest.mark.golden, pytest.mark.slow]
REAIS = Path(os.environ.get("COIR_NOTAS_REAIS", "")) if os.environ.get("COIR_NOTAS_REAIS") else None
ANONIMIZADAS = dados_dir() / "_notas-corretagens"


def _chave(n: NotaExtraida) -> tuple[object, ...]:
    return (n.corretora, n.numero, n.data_pregao, n.segmento)


def _conteudo(n: NotaExtraida) -> dict[str, object]:
    return n.model_dump(exclude={"conta", "arquivo", "paginas"})


@pytest.mark.skipif(REAIS is None or not REAIS.is_dir(), reason="COIR_NOTAS_REAIS ausente")
def test_extracao_das_reais_igual_a_das_anonimizadas() -> None:
    assert REAIS is not None
    reais, anonimas = extrair_pasta(REAIS), extrair_pasta(ANONIMIZADAS)
    assert [(r.arquivo, r.nota, r.motivo) for r in reais.rejeitadas] == []
    por_r = {_chave(n): n for n in reais.notas}
    por_a = {_chave(n): n for n in anonimas.notas}
    assert set(por_r) == set(por_a)
    diferentes = [k for k in por_r if _conteudo(por_r[k]) != _conteudo(por_a[k])]
    assert diferentes == []
    pares = {(por_r[k].corretora, por_r[k].conta, por_a[k].conta) for k in por_r}
    reais_por_conta = [(c, r) for c, r, _ in pares]
    assert len(reais_por_conta) == len(set(reais_por_conta)), "conta real em duas anonimizadas"
