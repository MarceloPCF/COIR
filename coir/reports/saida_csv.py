"""Relatório -> csv no padrão brasileiro (ADR 0003): `;`, vírgula decimal, sem separador de
milhar, data dd/mm/aaaa, UTF-8 com BOM (abre direto no Excel e no LibreOffice em português).

Seções com as mesmas colunas vão para um arquivo só, com a coluna "Seção" na frente; seções com
colunas diferentes geram um arquivo cada (`<chave>-<seção>.csv`). Só dados: o cabeçalho do
relatório (titular, CPF) e as linhas de total ficam no xlsx e no pdf; no csv, somar uma coluna
não conta o total duas vezes.
"""

import csv
from pathlib import Path

from coir.reports.modelo import Coluna, Relatorio, Secao, formatar, slug


def escrever_csv(rel: Relatorio, pasta: Path, prefixo: str) -> list[Path]:
    grupos: dict[tuple[Coluna, ...], list[Secao]] = {}
    for s in rel.secoes:
        grupos.setdefault(s.colunas, []).append(s)
    saida: list[Path] = []
    for colunas, secoes in grupos.items():
        sufixo = "" if len(grupos) == 1 else f"-{slug(secoes[0].titulo)}"
        caminho, n = pasta / f"{prefixo}-{rel.chave}{sufixo}.csv", 2
        while caminho in saida:  # duas seções com o mesmo nome curto
            caminho, n = pasta / f"{prefixo}-{rel.chave}{sufixo}-{n}.csv", n + 1
        com_secao = len(secoes) > 1
        with caminho.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=";", lineterminator="\r\n")
            w.writerow((["Seção"] if com_secao else []) + [c.nome for c in colunas])
            for s in secoes:
                for linha in s.linhas:
                    valores = [formatar(v, c.tipo, milhar=False)
                               for v, c in zip(linha, colunas, strict=True)]  # fmt: skip
                    w.writerow(([s.titulo] if com_secao else []) + valores)
        saida.append(caminho)
    return saida
