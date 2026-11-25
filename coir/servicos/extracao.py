"""Leitura dos PDFs de um lote em vários processos (importação).

Só a extração (texto do PDF, leitura e conciliação das notas) vai para os processos: é pura e
cada arquivo é independente. A conferência dos arquivos, a gravação (uma transação, um escritor
do SQLite) e a ordem do lote ficam no processo do programa: os resultados voltam na ordem dos
arquivos, e o lote dá o mesmo que a leitura um a um. O PDFium não é thread-safe (``TRAVA_PDFIUM``),
mas cada processo tem o seu.

Os processos nascem por ``spawn`` (o mesmo no Windows e no Linux; o programa empacotado chama
``multiprocessing.freeze_support`` nos dois pontos de entrada) e importam a leitura das notas e a
trilha da auditoria. A rejeição volta com o registro da auditoria já calculado; o evento
``erro.dominio`` de um erro criado dentro do processo de leitura (ex.: PDF ilegível) não vai para
o registro, que só existe no processo do programa: no lote em processos, esse erro aparece pela
rejeição do arquivo.
"""

import multiprocessing
import os
from collections.abc import Callable, Iterator, Sequence
from concurrent.futures import ProcessPoolExecutor
from concurrent.futures.process import BrokenProcessPool
from pathlib import Path

from coir.erros import ImportacaoError
from coir.parsing.extrair import ResultadoExtracao, extrair_pdf
from coir.servicos import trilha

MINIMO_PARALELO = 8  # abaixo disso, subir os processos custa mais que ler um a um


def _extrair(caminho: Path, senhas: tuple[str, ...]) -> ResultadoExtracao:
    """Roda no processo de leitura: a rejeição volta com os dados da auditoria já calculados
    (a exceção, com a pilha, não atravessa processos)."""
    resultado = extrair_pdf(caminho, senhas)
    for r in resultado.rejeitadas:
        r.auditoria = trilha.dados_rejeicao(r)
        r.erro = None
    return resultado


def processos(arquivos: int) -> int:
    """Quantos processos de leitura para ``arquivos`` PDFs: 1 (leitura um a um) abaixo do
    mínimo; senão um por núcleo, deixando um para a janela."""
    if arquivos < MINIMO_PARALELO:
        return 1
    return max(1, min(arquivos, (os.cpu_count() or 2) - 1))


Leitura = Callable[[Path, tuple[str, ...]], ResultadoExtracao]


def extrair_lote(
    caminhos: Sequence[Path], senhas: Sequence[str], ler: Leitura = _extrair
) -> Iterator[ResultadoExtracao]:
    """Resultados na ordem de ``caminhos`` (o próximo só sai depois do anterior). Fechado o
    gerador (fim do lote, erro na gravação), os processos param e o que não começou é cancelado.
    Um processo de leitura que cai (queda nativa ao ler um PDF) vira ``ImportacaoError``: o lote
    é desfeito, sem derrubar o programa. ``ler``: a leitura de cada arquivo (testes)."""
    n = processos(len(caminhos))
    if n == 1:
        yield from (extrair_pdf(c, senhas) for c in caminhos)
        return
    executor = ProcessPoolExecutor(max_workers=n, mp_context=multiprocessing.get_context("spawn"))
    try:
        yield from executor.map(ler, caminhos, [tuple(senhas)] * len(caminhos))
    except BrokenProcessPool as erro:
        raise ImportacaoError(
            "um processo de leitura dos PDFs parou de repente (um PDF pode estar danificado); "
            "nada foi gravado. Importe de novo em lotes menores (até 7 arquivos são lidos um a "
            "um) para achar o arquivo"
        ) from erro
    finally:
        executor.shutdown(wait=True, cancel_futures=True)
