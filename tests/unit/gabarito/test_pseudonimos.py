from pathlib import Path

from coir.gabarito.pseudonimos import MapaPseudonimos
from tests.unit.gabarito.fabrica import CPF_A, com_dv, mascara


def test_pseudonimos_estaveis_e_sequenciais(tmp_path: Path) -> None:
    mapa = MapaPseudonimos()
    assert mapa.titular(mascara(CPF_A)) == "TITULAR-1"
    assert mapa.titular(CPF_A) == "TITULAR-1"
    assert mapa.titular(com_dv("987654321")) == "TITULAR-2"
    assert mapa.conta("XP", "111") == "XP-1"
    assert mapa.conta("XP", "222") == "XP-2"
    assert mapa.conta("BTG", "111") == "BTG-1"
    arquivo = tmp_path / "_privado" / "mapa.json"
    mapa.salvar(arquivo)
    recarregado = MapaPseudonimos.carregar(arquivo)
    assert recarregado.titular(CPF_A) == "TITULAR-1"
    assert recarregado.conta("XP", "333") == "XP-3"
    assert set(recarregado.valores_reais()) >= {CPF_A, "111", "222"}


def test_carregar_inexistente_comeca_vazio(tmp_path: Path) -> None:
    assert MapaPseudonimos.carregar(tmp_path / "nao.json").titulares == {}


def test_repr_nao_expoe_valores_reais() -> None:
    mapa = MapaPseudonimos()
    mapa.titular(CPF_A)
    mapa.conta("XP", "987654")
    for texto in (repr(mapa), str(mapa), f"{mapa}"):
        assert CPF_A not in texto and "987654" not in texto
