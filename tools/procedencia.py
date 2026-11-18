"""Confere o manifesto de procedência das tabelas de referência (``data/procedencia.csv``).

Cada arquivo (ou pasta) de ``data/`` tem de estar no manifesto, com uma categoria conhecida, a
licença e a fonte. Tabela com a origem por linha (``POR_LINHA``) tem de ter a coluna de origem e
só origens da sua categoria. Arquivo novo sem procedência, ou entrada do manifesto sem arquivo,
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
CATEGORIAS = frozenset({"projeto", "curadoria", "norma", "cvm", "bcb", "b3"})
# tabelas com a origem de cada linha (coluna origem; "a+b" é a soma de fontes): as origens têm de
# ser da categoria do arquivo
POR_LINHA = {"eventos_cvm.csv": "origem", "eventos_b3.csv": "origem"}
ORIGENS_DA_CATEGORIA = {"cvm": frozenset({"cvm_fre", "cvm_ipe"}), "b3": frozenset({"b3"})}


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


def origens(pasta: Path, arquivo: str, categoria: str) -> Counter[str]:
    """Quantas linhas de cada origem a tabela tem; origem fora da categoria falha."""
    if arquivo not in POR_LINHA:
        raise ProcedenciaError(f"{arquivo}: sem coluna de origem declarada")
    coluna, aceitas = POR_LINHA[arquivo], ORIGENS_DA_CATEGORIA.get(categoria, frozenset())
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
    for nome in POR_LINHA:
        origens(pasta, nome, entradas[nome].categoria)
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
    for nome in sorted(POR_LINHA):
        contagem = origens(DATA, nome, entradas[nome].categoria)
        resumo = ", ".join(f"{o} {n}" for o, n in sorted(contagem.items()))
        print(f"{nome}: {resumo}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
