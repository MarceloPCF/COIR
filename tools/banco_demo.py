"""Banco de demonstração para a interface, com titular fictício (CPF sintético); nenhum dado real.

- Padrão (testes da UI): as notas anonimizadas de tests/fixtures/notas. Com ``pdfs=True``, cada
  nota vira também um PDF (o texto da fixture) em ``<pasta>/_notas-corretagens``, para a tela de
  notas mostrar o original ao lado.
- Com ``roteiro`` (capturas de tela e manual): as notas fictícias do gerador
  (tools/notas_ficticias.py), em PDF, lidas pelo extrator real, e os proventos do roteiro pela
  planilha de movimentação da B3.

Uso: uv run python -m tools.banco_demo <pasta> [--fixtures]
     (cria <pasta>/coir.db e os PDFs; sem --fixtures, com o roteiro de demonstração)
"""

import sys
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

from coir import referencias
from coir.db import repositorio as repo
from coir.db.sessao import NOME_BANCO, conectar, migrar
from coir.parsing.extrair import ResultadoExtracao, extrair_textos
from coir.servicos.b3 import importar_planilhas
from coir.servicos.darf import registrar_pagamento
from coir.servicos.estado import Estado
from coir.servicos.importacao import importar_arquivos
from coir.servicos.notas import PASTA_NOTAS
from tests.apoio import cpf_de_teste
from tools import notas_ficticias

FIXTURES = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "notas"
AGORA = datetime(2026, 10, 16, 12, 0, 0)
NOME = "Titular de Demonstração"


def extrator(cpf: str) -> Callable[[Path], ResultadoExtracao]:
    """Lê o texto da fixture (o próprio arquivo, ou a fixture de mesmo nome se for o PDF da
    demonstração) com o CPF anonimizado trocado pelo de teste."""
    mascara = f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}"

    def extrair(caminho: Path) -> ResultadoExtracao:
        fonte = FIXTURES / f"{caminho.stem}.txt" if caminho.suffix == ".pdf" else caminho
        if caminho.suffix == ".pdf" and not caminho.read_bytes().startswith(b"%PDF"):
            fonte = caminho  # texto enviado com nome .pdf (testes da tela de importação)
        texto = fonte.read_text(encoding="utf-8")
        texto = texto.replace("000.000.000-00", f"{mascara}-{cpf[9:]}").replace(
            "000.000.000.00", f"{mascara}.{cpf[9:]}"
        )
        return extrair_textos(caminho.name, texto.split("\f"))

    return extrair


def notas() -> list[Path]:
    """As fixtures, menos o exercício de opção: a série exercida não está nas outras notas e o
    motor falha alto com histórico incompleto."""
    return [p for p in sorted(FIXTURES.glob("*.txt")) if "exercicio" not in p.stem]


def criar(pasta: Path, pdfs: bool = False, roteiro: Path | None = None) -> Path:
    banco = pasta / NOME_BANCO
    pasta.mkdir(parents=True, exist_ok=True)
    migrar(banco, AGORA)
    if roteiro is not None:
        return _do_roteiro(banco, notas_ficticias.carregar(roteiro))
    cpf = cpf_de_teste()
    arquivos = notas()
    if pdfs:
        destino = pasta / PASTA_NOTAS
        destino.mkdir(exist_ok=True)
        arquivos = [notas_ficticias.pdf(a.read_text(encoding="utf-8"), destino / f"{a.stem}.pdf")
                    for a in arquivos]  # fmt: skip
    with conectar(banco) as fabrica:
        with fabrica.begin() as s:
            if repo.titular_por_cpf(s, cpf) is None:
                repo.adicionar_titular(s, cpf, NOME)
        with fabrica() as s:
            importar_arquivos(s, arquivos, referencias.carregar(), AGORA,
                              extrator=extrator(cpf))  # fmt: skip
    return banco


def _do_roteiro(banco: Path, roteiro: notas_ficticias.Roteiro) -> Path:
    pasta = banco.parent
    arquivos, planilha = notas_ficticias.gerar(roteiro, pasta / PASTA_NOTAS)
    with conectar(banco) as fabrica:
        with fabrica.begin() as s:
            titular = repo.titular_por_cpf(s, roteiro.cpf)
            if titular is None:
                titular = repo.adicionar_titular(s, roteiro.cpf, roteiro.titular)
            id_titular = titular.id
        with fabrica() as s:
            rel = importar_arquivos(s, arquivos, referencias.carregar(), AGORA)
    if rel.rejeitadas:  # o roteiro gerou nota que não fecha: banco e capturas sairiam incompletos
        raise RuntimeError(
            "notas fictícias rejeitadas: " + "; ".join(r.motivo for r in rel.rejeitadas)
        )
    if planilha is not None:
        destino = pasta / "planilhas-b3"
        destino.mkdir(exist_ok=True)
        importar_planilhas(banco, [planilha.replace(destino / planilha.name)], id_titular, AGORA)
    if roteiro.darfs_pagas:
        estado = Estado(banco)
        estado.preparar(AGORA)
        estado.titular = id_titular
        for linha in estado.darfs_todos_os_anos():
            pago_em = roteiro.darfs_pagas.get((linha.ano, linha.mes))
            if pago_em is not None:
                registrar_pagamento(banco, pasta, id_titular, linha.ano, linha.mes, pago_em,
                                    linha.principal, AGORA)  # fmt: skip
    return banco


if __name__ == "__main__":
    fixtures = "--fixtures" in sys.argv[2:]
    roteiro_demo = None if fixtures else notas_ficticias.ROTEIRO_DEMONSTRACAO
    print(criar(Path(sys.argv[1]), pdfs=True, roteiro=roteiro_demo))
