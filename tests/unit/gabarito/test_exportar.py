import json
from pathlib import Path

import pytest

from coir.erros import GabaritoError
from coir.gabarito.exportar import (
    ARQUIVO_MANIFESTO,
    ARQUIVOS,
    RAIZ_REPOSITORIO,
    carregar,
    exportar,
    verificar_pseudonimizacao,
)
from coir.gabarito.extrair import extrair
from coir.gabarito.pseudonimos import MapaPseudonimos
from tests.unit.gabarito import fabrica
from tools.verificar_dados_pessoais import verificar


def _exportar(destino: Path) -> tuple[dict[str, object], MapaPseudonimos]:
    destino.parent.mkdir(parents=True, exist_ok=True)
    origem = destino.parent / "gabarito.xlsb"
    origem.write_bytes(b"planilha ficticia")
    mapa = MapaPseudonimos()
    gabarito = extrair(fabrica.planilha(), mapa)
    return exportar(gabarito, mapa, origem, destino), mapa


def test_exporta_e_recarrega_sem_perdas(tmp_path: Path) -> None:
    destino = tmp_path / "gabarito"
    manifesto, _ = _exportar(destino)
    original = extrair(fabrica.planilha(), MapaPseudonimos())
    assert carregar(destino) == original
    assert manifesto["contagens"]["operacoes"] == {"COMUM": 2, "DAY_TRADE": 1}  # type: ignore[index]
    assert set(json.loads((destino / ARQUIVO_MANIFESTO).read_text("utf-8"))["arquivos"]) == {
        nome for nome, _ in ARQUIVOS.values()
    }


def test_exportacao_deterministica(tmp_path: Path) -> None:
    m1, _ = _exportar(tmp_path / "a" / "gabarito")
    m2, _ = _exportar(tmp_path / "b" / "gabarito")
    assert m1["arquivos"] == m2["arquivos"]


def test_arquivos_publicos_sem_cpf_nem_conta(tmp_path: Path) -> None:
    destino = tmp_path / "gabarito"
    _, mapa = _exportar(destino)
    termos = [v.casefold() for v in mapa.valores_reais()]
    publicos = [destino / nome for nome, _ in ARQUIVOS.values()] + [destino / ARQUIVO_MANIFESTO]
    assert verificar(publicos, termos=termos) == []
    assert (destino / "_privado" / "mapa-pseudonimos.json").is_file()


def test_verificacao_acusa_conta_real(tmp_path: Path) -> None:
    destino = tmp_path / "gabarito"
    _, mapa = _exportar(destino)
    csv = destino / "operacoes.csv"
    texto = csv.read_text("utf-8").replace("AGORA-1", "111222", 1)
    csv.write_text(texto, "utf-8")
    with pytest.raises(GabaritoError, match="valor real"):
        verificar_pseudonimizacao(destino, mapa)


def test_recusa_gravar_dentro_do_repositorio(tmp_path: Path) -> None:
    origem = tmp_path / "gabarito.xlsb"
    origem.write_bytes(b"x")
    mapa = MapaPseudonimos()
    gabarito = extrair(fabrica.planilha(), mapa)
    with pytest.raises(GabaritoError, match="repositório"):
        exportar(gabarito, mapa, origem, RAIZ_REPOSITORIO / "gabarito-teste")
    assert not (RAIZ_REPOSITORIO / "gabarito-teste").exists()


def test_carregar_arquivo_ausente(tmp_path: Path) -> None:
    with pytest.raises(GabaritoError, match="ausente"):
        carregar(tmp_path)
