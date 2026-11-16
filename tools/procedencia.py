"""Confere o manifesto de procedência das tabelas de referência (``data/procedencia.csv``).

Cada arquivo (ou pasta) de ``data/`` tem de estar no manifesto, com uma categoria conhecida, a
licença e a fonte. Tabela com a origem por linha (``por_linha``) tem de ter a coluna de origem e
só valores conhecidos nela. Arquivo novo sem procedência, ou entrada do manifesto sem arquivo,
falha alto.

Uso: uv run python -m tools.procedencia
"""

import csv
import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
DATA = RAIZ / "data"
MANIFESTO = "procedencia.csv"
COLUNAS = ("arquivo", "categoria", "licenca", "fonte")
CATEGORIAS = frozenset({"projeto", "curadoria", "norma", "cvm", "bcb", "b3", "por_linha"})
# tabelas com origem por linha: arquivo -> (coluna, origens aceitas; "a+b" é a soma de fontes)
POR_LINHA = {"eventos_b3.csv": ("origem", frozenset({"b3", "cvm_fre", "cvm_ipe"}))}


class ProcedenciaError(Exception):
    """Arquivo de data/ sem procedência declarada, ou manifesto inconsistente."""


@dataclass(frozen=True)
class Entrada:
    arquivo: str
    categoria: str
    licenca: str
    fonte: str


def ler(pasta: Path = DATA) -> dict[str, Entrada]:
    caminho = pasta / MANIFESTO
    with caminho.open(encoding="utf-8", newline="") as f:
        leitor = csv.DictReader(f)
        if tuple(leitor.fieldnames or ()) != COLUNAS:
            raise ProcedenciaError(f"{MANIFESTO}: colunas devem ser {', '.join(COLUNAS)}")
        entradas: dict[str, Entrada] = {}
        for r in leitor:
            e = Entrada(**{c: (r[c] or "").strip() for c in COLUNAS})
            if e.arquivo in entradas:
                raise ProcedenciaError(f"{MANIFESTO}: {e.arquivo} repetido")
            if e.categoria not in CATEGORIAS:
                raise ProcedenciaError(f"{MANIFESTO}: categoria desconhecida em {e.arquivo}")
            if not e.licenca or not e.fonte:
                raise ProcedenciaError(f"{MANIFESTO}: {e.arquivo} sem licença ou sem fonte")
            entradas[e.arquivo] = e
    return entradas


def origens(pasta: Path, arquivo: str) -> Counter[str]:
    """Quantas linhas de cada origem a tabela ``por_linha`` tem; origem desconhecida falha."""
    if arquivo not in POR_LINHA:
        raise ProcedenciaError(f"{arquivo}: sem coluna de origem declarada")
    coluna, aceitas = POR_LINHA[arquivo]
    with (pasta / arquivo).open(encoding="utf-8", newline="") as f:
        linhas = list(csv.DictReader(f))
    if linhas and coluna not in linhas[0]:
        raise ProcedenciaError(f"{arquivo}: falta a coluna {coluna}")
    contagem: Counter[str] = Counter()
    for r in linhas:
        partes = (r[coluna] or "").split("+")
        if not all(p in aceitas for p in partes):
            raise ProcedenciaError(f"{arquivo}: origem desconhecida {r[coluna]!r}")
        contagem.update(partes)
    return contagem


def conferir(pasta: Path = DATA) -> dict[str, Entrada]:
    """Manifesto e pasta batem um com o outro; devolve o manifesto."""
    entradas = ler(pasta)
    presentes = {p.name for p in pasta.iterdir() if not p.name.startswith((".", "__"))}
    sem_procedencia = sorted(presentes - set(entradas))
    if sem_procedencia:
        raise ProcedenciaError(f"sem procedência em {MANIFESTO}: {', '.join(sem_procedencia)}")
    sem_arquivo = sorted(set(entradas) - presentes)
    if sem_arquivo:
        raise ProcedenciaError(f"{MANIFESTO} cita o que não existe: {', '.join(sem_arquivo)}")
    for nome, e in entradas.items():
        if e.categoria == "por_linha":
            origens(pasta, nome)
        elif nome in POR_LINHA:
            raise ProcedenciaError(f"{nome}: tem origem por linha, categoria deve ser por_linha")
    return entradas


def main() -> int:
    try:
        entradas = conferir()
    except ProcedenciaError as erro:
        print(f"ERRO: {erro}", file=sys.stderr)
        return 1
    por_categoria = Counter(e.categoria for e in entradas.values())
    for categoria, n in sorted(por_categoria.items()):
        print(f"{categoria}: {n}")
    for nome in sorted(n for n, e in entradas.items() if e.categoria == "por_linha"):
        resumo = ", ".join(f"{o} {n}" for o, n in sorted(origens(DATA, nome).items()))
        print(f"{nome}: {resumo}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
