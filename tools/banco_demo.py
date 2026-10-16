"""Banco de demonstração para a interface: as notas anonimizadas de tests/fixtures/notas com um
titular fictício (CPF sintético). Serve às capturas de tela e aos testes da UI; nenhum dado real.

Uso: uv run python -m tools.banco_demo <pasta>   (cria <pasta>/coir.db)
"""

import sys
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

from coir import referencias
from coir.db import repositorio as repo
from coir.db.sessao import NOME_BANCO, conectar, migrar
from coir.parsing.extrair import ResultadoExtracao, extrair_textos
from coir.servicos.importacao import importar_arquivos
from tests.apoio import cpf_de_teste

FIXTURES = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "notas"
AGORA = datetime(2026, 10, 16, 12, 0, 0)
NOME = "Titular de Demonstração"


def _extrator(cpf: str) -> Callable[[Path], ResultadoExtracao]:
    mascara = f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}"

    def extrair(caminho: Path) -> ResultadoExtracao:
        texto = caminho.read_text(encoding="utf-8")
        texto = texto.replace("000.000.000-00", f"{mascara}-{cpf[9:]}").replace(
            "000.000.000.00", f"{mascara}.{cpf[9:]}"
        )
        return extrair_textos(caminho.name, texto.split("\f"))

    return extrair


def _notas() -> list[Path]:
    """As fixtures, menos o exercício de opção: a série exercida não está nas outras notas e o
    motor falha alto com histórico incompleto."""
    return [p for p in sorted(FIXTURES.glob("*.txt")) if "exercicio" not in p.stem]


def criar(pasta: Path) -> Path:
    banco = pasta / NOME_BANCO
    pasta.mkdir(parents=True, exist_ok=True)
    migrar(banco, AGORA)
    cpf = cpf_de_teste()
    with conectar(banco) as fabrica:
        with fabrica.begin() as s:
            if repo.titular_por_cpf(s, cpf) is None:
                repo.adicionar_titular(s, cpf, NOME)
        with fabrica() as s:
            importar_arquivos(s, _notas(), referencias.carregar(), AGORA,
                              extrator=_extrator(cpf))  # fmt: skip
    return banco


if __name__ == "__main__":
    print(criar(Path(sys.argv[1])))
