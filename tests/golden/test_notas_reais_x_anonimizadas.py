"""A base anonimizada de desenvolvimento (COIR_DADOS/_notas-anonimizadas) só vale como base de
teste se a extração dela for a mesma das notas originais. Na 3.0.1 não era: as cópias antigas
(_notas-corretagens) foram anonimizadas mexendo no PDF, isso levou a conta da BTG para a última
linha e o leitor, ajustado às cópias, lia nas originais o telefone da ouvidoria. A base nova é o
texto das originais com os dados pessoais trocados no lugar (``coir notas anonimizar``).

Roda com COIR_NOTAS_REAIS apontando para a pasta das notas originais (fora do repositório; só no
PC do mantenedor). Compara nota a nota: as mesmas notas, nada rejeitado e conteúdo igual fora a
conta; cada conta real corresponde a uma só conta anonimizada, a do mapa de anonimização."""

import os
from pathlib import Path

import pytest

from coir.config import dados_dir
from coir.domain.nota import NotaExtraida
from coir.parsing.extrair import extrair_pasta
from coir.servicos.suporte import ARQUIVO_MAPA, carregar_mapa
from tests.golden.conftest import NOTAS_ANONIMIZADAS, SUFIXO_NOTAS

pytestmark = [pytest.mark.golden, pytest.mark.slow]
REAIS = Path(os.environ.get("COIR_NOTAS_REAIS", "")) if os.environ.get("COIR_NOTAS_REAIS") else None


def _chave(n: NotaExtraida) -> tuple[object, ...]:
    return (n.corretora, n.numero, n.data_pregao, n.segmento)


def _conteudo(n: NotaExtraida) -> dict[str, object]:
    return n.model_dump(exclude={"conta", "arquivo"})


@pytest.mark.skipif(REAIS is None or not REAIS.is_dir(), reason="COIR_NOTAS_REAIS ausente")
def test_extracao_das_reais_igual_a_das_anonimizadas() -> None:
    assert REAIS is not None
    reais, anonimas = extrair_pasta(REAIS), extrair_pasta(NOTAS_ANONIMIZADAS, SUFIXO_NOTAS)
    assert [(r.arquivo, r.nota, r.motivo) for r in reais.rejeitadas] == []
    assert sorted(p.stem for p in REAIS.glob("*.pdf")) == sorted(
        p.stem for p in NOTAS_ANONIMIZADAS.glob(f"*{SUFIXO_NOTAS}")
    )
    por_r = {_chave(n): n for n in reais.notas}
    por_a = {_chave(n): n for n in anonimas.notas}
    assert set(por_r) == set(por_a)
    diferentes = [k for k in por_r if _conteudo(por_r[k]) != _conteudo(por_a[k])]
    assert diferentes == []
    assert [k for k in por_r if Path(por_r[k].arquivo).stem != Path(por_a[k].arquivo).stem] == []
    mapa = carregar_mapa(dados_dir() / ARQUIVO_MAPA)
    ficticia = {k: mapa.contas[f"{n.corretora}|{n.conta}"] for k, n in por_r.items()}
    trocadas = [k for k in por_r if por_a[k].conta != ficticia[k]]
    assert trocadas == [], "conta anonimizada diferente da do mapa"
    assert not {n.conta for n in reais.notas} & {n.conta for n in anonimas.notas}
