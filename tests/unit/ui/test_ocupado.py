"""Cortina "Processando dados, aguarde": aparece enquanto o COIR grava e some no fim, mesmo com
erro, para ninguém trocar de tela no meio de uma importação."""

import pytest

from coir.ui import ocupado


def test_processando_marca_e_desmarca_a_situacao() -> None:
    assert not ocupado.SITUACAO.ativo
    with ocupado.processando("Importando as notas"):
        assert ocupado.SITUACAO.ativo
        assert ocupado.SITUACAO.texto == "Importando as notas"
        ocupado.detalhar("3 de 12 arquivos")
        assert ocupado.SITUACAO.detalhe == "3 de 12 arquivos"
    assert not ocupado.SITUACAO.ativo
    assert ocupado.SITUACAO.texto == "" and ocupado.SITUACAO.detalhe == ""


def test_erro_no_bloco_nao_deixa_a_cortina_presa() -> None:
    with pytest.raises(ValueError, match="falhou"), ocupado.processando("Restaurando o backup"):
        raise ValueError("falhou")
    assert not ocupado.SITUACAO.ativo


def test_operacoes_aninhadas_so_liberam_no_fim_da_ultima() -> None:
    with ocupado.processando("Atualizando as tabelas"):
        with ocupado.processando("Conferindo os eventos"):
            assert ocupado.SITUACAO.operacoes == 2
            assert ocupado.SITUACAO.texto == "Conferindo os eventos"
        assert ocupado.SITUACAO.ativo
        assert ocupado.SITUACAO.texto == "Atualizando as tabelas"  # volta o da operação de fora
    assert not ocupado.SITUACAO.ativo
