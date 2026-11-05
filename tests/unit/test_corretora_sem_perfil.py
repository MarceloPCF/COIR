"""Corretora sem perfil do PDF ao banco (pedido do usuário de 05/11/2026): a nota fictícia lida pelo
layout padrão é gravada com o código da B3 como corretora, e a importação avisa o cliente."""

from datetime import datetime
from pathlib import Path

from coir import referencias
from coir.servicos.importacao import importar_lista
from tools import notas_ficticias


def test_nota_de_corretora_sem_perfil_entra_com_aviso(tmp_path: Path) -> None:
    _, texto = notas_ficticias.textos(notas_ficticias.carregar())[0]
    nova = (
        texto.replace("XP INVESTIMENTOS", "CORRETORA NOVA")
        .replace("02.332.886/0001-04", "00.000.000/0001-91")
        .replace("3-5 ", "999-9 ")
    )
    assert nova != texto and "XP INVESTIMENTOS" not in nova
    pdf = notas_ficticias.pdf(nova, tmp_path / "nova.pdf")
    rel = importar_lista(tmp_path / "coir.db", [pdf], referencias.carregar(),
                         datetime(2026, 11, 5, 12))  # fmt: skip
    assert rel.notas_gravadas == 1 and not rel.rejeitadas
    assert any("nova.pdf" in a and "lida pelo layout padrão" in a for a in rel.avisos)
