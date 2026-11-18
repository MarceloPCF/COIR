"""Manifesto de procedência de data/: todo arquivo declarado, com categoria, licença e fonte."""

import csv
import shutil
from pathlib import Path

import pytest

from tools.procedencia import DATA, ProcedenciaError, conferir, origens


def test_data_real_tem_procedencia_completa() -> None:
    entradas = conferir()
    curadas = ("instituicoes_b3.csv",)  # "b3" só no nome
    assert all(e.categoria == "b3" for n, e in entradas.items()
               if n.endswith("_b3.csv") and n not in curadas)  # fmt: skip
    assert entradas["selic_mensal.csv"].licenca == "ODbL-1.0"


def test_eventos_tem_so_origens_da_sua_fonte() -> None:
    assert set(origens(DATA, "eventos_cvm.csv", "cvm")) == {"cvm_fre", "cvm_ipe"}
    assert set(origens(DATA, "eventos_b3.csv", "b3")) == {"b3"}


def _copia(tmp_path: Path) -> Path:
    pasta = tmp_path / "data"
    shutil.copytree(DATA, pasta)
    return pasta


def test_arquivo_novo_sem_procedencia_falha(tmp_path: Path) -> None:
    pasta = _copia(tmp_path)
    (pasta / "novo.csv").write_text("a\n1\n", encoding="utf-8")
    with pytest.raises(ProcedenciaError, match=r"novo\.csv"):
        conferir(pasta)


def test_entrada_sem_arquivo_falha(tmp_path: Path) -> None:
    pasta = _copia(tmp_path)
    (pasta / "feriados.csv").unlink()
    with pytest.raises(ProcedenciaError, match=r"feriados\.csv"):
        conferir(pasta)


def test_categoria_desconhecida_falha(tmp_path: Path) -> None:
    pasta = _copia(tmp_path)
    manifesto = pasta / "procedencia.csv"
    texto = manifesto.read_text(encoding="utf-8")
    manifesto.write_text(texto.replace("selic_mensal.csv,bcb", "selic_mensal.csv,x"), "utf-8")
    with pytest.raises(ProcedenciaError, match="categoria"):
        conferir(pasta)


def test_origem_desconhecida_falha(tmp_path: Path) -> None:
    pasta = _copia(tmp_path)
    eventos = pasta / "eventos_cvm.csv"
    with eventos.open(encoding="utf-8", newline="") as f:
        linhas = list(csv.DictReader(f))
    linhas[0]["origem"] = "cvm_fre+b3"  # linha da bolsa no arquivo da CVM
    with eventos.open("w", encoding="utf-8", newline="") as f:
        escritor = csv.DictWriter(f, fieldnames=list(linhas[0]))
        escritor.writeheader()
        escritor.writerows(linhas)
    with pytest.raises(ProcedenciaError, match="origem desconhecida"):
        conferir(pasta)
