"""Eventos lançados por arquivo CSV (o mesmo formato da exportação do gabarito)."""

import csv
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from pathlib import Path

from sqlalchemy import select
from sqlalchemy.orm import Session

from coir.db import repositorio as repo
from coir.db.modelos import (
    Conta,
    Corretora,
    EventoCorporativo,
)
from coir.domain.fiscal import (
    NaturezaEvento,
    TipoEvento,
)
from coir.erros import ImportacaoError
from coir.servicos.eventos.dados import (
    COLUNAS,
    DadosEvento,
    _chaves,
    _decimal,
    _preencher,
    _validar,
)


@dataclass
class RelatorioEventos:
    novos: int = 0
    ja_no_banco: int = 0
    fora_do_csv: int = 0  # no banco e não no CSV (linha corrigida ou apagada)
    removidos: int = 0


def _conta(sessao: Session, corretora: str, numero: str, linha: int) -> Conta:
    conta = sessao.scalar(
        select(Conta).join(Corretora).where(Corretora.nome == corretora, Conta.numero == numero)
    )
    if conta is None:
        raise ImportacaoError(
            f"eventos linha {linha}: conta da {corretora} sem notas no banco (importe as notas)"
        )
    return conta


def _ativo(sessao: Session, ticker: str, dia: date, linha: int) -> int:
    """Ativo já negociado (evento de ticker sem nota no banco é erro de digitação provável)."""
    ativo = repo.ativo_existente(sessao, ticker, dia) or repo.ativo_pelo_codigo_antigo(
        sessao, ticker, dia
    )
    if ativo is None:
        raise ImportacaoError(f"eventos linha {linha}: {ticker} sem negócio no banco em {dia}")
    return ativo


ORIGENS_TELA = frozenset({"tela", "b3"})  # lançados na tela 7: o CSV com --substituir não apaga


def ler_csv(caminho: Path) -> list[dict[str, str]]:
    """CSV em UTF-8 (com ou sem BOM), separado por vírgula ou ponto e vírgula (Excel pt-BR)."""
    try:
        with caminho.open(encoding="utf-8-sig", newline="") as arquivo:
            cabecalho = arquivo.readline()
            arquivo.seek(0)
            leitor = csv.DictReader(
                arquivo, delimiter=";" if cabecalho.count(";") > cabecalho.count(",") else ","
            )
            faltam = [c for c in COLUNAS if c not in (leitor.fieldnames or [])]
            if faltam:
                raise ImportacaoError(f"{caminho.name}: colunas ausentes: {', '.join(faltam)}")
            return [{k: (v or "").strip() for k, v in linha.items()} for linha in leitor]
    except OSError as erro:
        raise ImportacaoError(f"não foi possível ler {caminho}: {erro.strerror or erro}") from erro
    except UnicodeDecodeError as erro:
        raise ImportacaoError(
            f'{caminho.name}: não é um CSV em UTF-8 (salve como "CSV UTF-8" no Excel)'
        ) from erro


def importar_eventos(
    sessao: Session, linhas: list[dict[str, str]], substituir: bool = False
) -> RelatorioEventos:
    """Grava as linhas novas numa transação (qualquer linha inválida desfaz tudo). Eventos do
    banco que não estão no CSV são contados; com `substituir`, removidos (o CSV é a fonte), menos
    os lançados na tela 7."""
    rel = RelatorioEventos()
    vistas: dict[str, int] = {}
    with sessao.begin():
        for n, linha in enumerate(linhas, start=2):
            chave, outra = _chaves(linha, n)
            if chave in vistas or outra in vistas:
                anterior = vistas.get(chave, vistas.get(outra))
                raise ImportacaoError(f"eventos linha {n}: repete a linha {anterior}")
            vistas[chave] = vistas[outra] = n
            if sessao.scalar(select(EventoCorporativo).where(
                    EventoCorporativo.chave.in_((chave, outra)))):  # fmt: skip
                rel.ja_no_banco += 1
                continue
            try:
                tipo = TipoEvento(linha["tipo"].upper())
                dia = date.fromisoformat(linha["data"])
                natureza = (NaturezaEvento(linha["natureza"].upper()) if linha.get("natureza")
                            else None)  # fmt: skip
            except ValueError as erro:
                raise ImportacaoError(f"eventos linha {n}: {erro}") from erro
            destino = None
            if tipo is TipoEvento.TRANSFERENCIA:
                destino = _conta(sessao, linha["corretora_destino"], linha["conta_destino"], n)
            conta = _conta(sessao, linha["corretora"], linha["conta"], n)
            d = DadosEvento(
                tipo=tipo, data=dia, conta=conta.id, ticker=linha["ticker"].upper(),
                cv=linha["cv"].upper() or None,
                quantidade=_decimal(linha["quantidade"], "quantidade", n),
                valor=_decimal(linha["valor"], "valor", n) or Decimal(0),
                conta_destino=None if destino is None else destino.id,
                ticker_destino=linha.get("ticker_destino", "").upper() or None,
                quantidade_destino=_decimal(linha.get("quantidade_destino", ""),
                                            "quantidade_destino", n),
                natureza=natureza,
                percentual_custo=_decimal(linha.get("percentual_custo", ""), "percentual_custo",
                                          n),
                valor_corrente=_decimal(linha.get("valor_corrente", ""), "valor_corrente", n),
                observacao=linha["observacao"], origem=linha["origem"] or "manual",
            )  # fmt: skip
            try:
                _validar(d)
            except ImportacaoError as erro:
                raise ImportacaoError(f"eventos linha {n}: {erro}") from erro
            if destino is not None and destino.titular_id != conta.titular_id:
                raise ImportacaoError(
                    f"eventos linha {n}: transferência entre titulares diferentes"
                )
            registro = EventoCorporativo()
            _preencher(sessao, registro, d, chave, _ativo(sessao, d.ticker, dia, n))
            sessao.add(registro)
            rel.novos += 1
        fora = [e for e in sessao.scalars(select(EventoCorporativo))
                if e.chave not in vistas and e.origem not in ORIGENS_TELA]  # fmt: skip
        if substituir:
            for e in fora:
                sessao.delete(e)
            rel.removidos = len(fora)
        else:
            rel.fora_do_csv = len(fora)
    return rel
