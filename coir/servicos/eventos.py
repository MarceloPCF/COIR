"""Eventos corporativos e de custódia: CSV mantido pelo usuário (em COIR_DADOS) -> banco.

Colunas: data, corretora, conta, ticker, tipo, cv, quantidade, valor, corretora_destino,
conta_destino, observacao, origem. Ajuste de posição: C soma quantidade e custo (valor), V
subtrai (como na aba Eventos_B3 do gabarito). Transferência: cv vazio, conta de destino e
quantidade (vazia = toda a posição). Reimportar o mesmo CSV não duplica (chave = hash da linha).
"""

import csv
import hashlib
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path

from sqlalchemy import select
from sqlalchemy.orm import Session

from coir.db import repositorio as repo
from coir.db.modelos import Conta, Corretora, EventoCorporativo
from coir.domain.fiscal import TipoEvento
from coir.erros import ImportacaoError

COLUNAS = ("data", "corretora", "conta", "ticker", "tipo", "cv", "quantidade", "valor",
           "corretora_destino", "conta_destino", "observacao", "origem")  # fmt: skip


@dataclass
class RelatorioEventos:
    novos: int = 0
    ja_no_banco: int = 0
    fora_do_csv: int = 0  # no banco e não no CSV (linha corrigida ou apagada)
    removidos: int = 0


def _decimal(texto: str, campo: str, linha: int) -> Decimal | None:
    if not texto:
        return None
    try:
        valor = Decimal(texto)
    except InvalidOperation as erro:
        raise ImportacaoError(f"eventos linha {linha}: {campo} inválido ({texto!r})") from erro
    if not valor.is_finite() or valor < 0:  # o sentido vem da coluna C/V, nunca do sinal
        raise ImportacaoError(f"eventos linha {linha}: {campo} deve ser número >= 0 ({texto!r})")
    return valor


EXIGE_QUANTIDADE = {TipoEvento.DESDOBRAMENTO, TipoEvento.GRUPAMENTO, TipoEvento.BONIFICACAO,
                    TipoEvento.SUBSCRICAO}  # fmt: skip


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
    ativo = repo.ativo_existente(sessao, ticker, dia)
    if ativo is None:
        raise ImportacaoError(f"eventos linha {linha}: {ticker} sem negócio no banco em {dia}")
    return ativo


def _chave(linha: dict[str, str], n: int) -> str:
    """Hash dos campos que definem o evento, normalizados ("100" = "100.0"; tipo e C/V em
    maiúsculas); observação e origem não entram."""
    campos = [linha.get(c, "").upper() for c in COLUNAS[:10]]
    for i in (6, 7):  # quantidade, valor
        numero = _decimal(campos[i], COLUNAS[i], n)
        campos[i] = "" if numero is None else format(numero.normalize(), "f")
    return hashlib.sha256("|".join(campos).encode()).hexdigest()


def ler_csv(caminho: Path) -> list[dict[str, str]]:
    """CSV em UTF-8 (com ou sem BOM), separado por vírgula ou ponto e vírgula (Excel pt-BR)."""
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


def importar_eventos(
    sessao: Session, linhas: list[dict[str, str]], substituir: bool = False
) -> RelatorioEventos:
    """Grava as linhas novas numa transação (qualquer linha inválida desfaz tudo). Eventos do
    banco que não estão no CSV são contados; com `substituir`, removidos (o CSV é a fonte)."""
    rel = RelatorioEventos()
    vistas: dict[str, int] = {}
    with sessao.begin():
        for n, linha in enumerate(linhas, start=2):
            chave = _chave(linha, n)
            if chave in vistas:
                raise ImportacaoError(f"eventos linha {n}: repete a linha {vistas[chave]}")
            vistas[chave] = n
            if sessao.scalar(select(EventoCorporativo).where(EventoCorporativo.chave == chave)):
                rel.ja_no_banco += 1
                continue
            try:
                tipo = TipoEvento(linha["tipo"].upper())
                dia = date.fromisoformat(linha["data"])
            except ValueError as erro:
                raise ImportacaoError(f"eventos linha {n}: {erro}") from erro
            cv = linha["cv"].upper() or None
            if (cv is None) != (tipo is TipoEvento.TRANSFERENCIA) or cv not in (None, "C", "V"):
                raise ImportacaoError(
                    f"eventos linha {n}: C/V {linha['cv']!r} inválido para {tipo}"
                )
            destino = None
            if tipo is TipoEvento.TRANSFERENCIA:
                destino = _conta(sessao, linha["corretora_destino"], linha["conta_destino"], n)
            conta = _conta(sessao, linha["corretora"], linha["conta"], n)
            quantidade = _decimal(linha["quantidade"], "quantidade", n)
            valor = _decimal(linha["valor"], "valor", n) or Decimal(0)
            if tipo in EXIGE_QUANTIDADE and not quantidade:
                raise ImportacaoError(f"eventos linha {n}: {tipo} sem quantidade")
            if tipo is TipoEvento.REDUCAO_CAPITAL and not valor:
                raise ImportacaoError(f"eventos linha {n}: redução de capital sem valor")
            if destino is not None and destino.titular_id != conta.titular_id:
                raise ImportacaoError(
                    f"eventos linha {n}: transferência entre titulares diferentes"
                )
            sessao.add(
                EventoCorporativo(
                    chave=chave, data=dia, conta=conta,
                    ativo_id=_ativo(sessao, linha["ticker"].upper(), dia, n), tipo=tipo.value,
                    cv=cv, quantidade=quantidade, valor=valor,
                    conta_destino=destino, observacao=linha["observacao"],
                    origem=linha["origem"] or "manual",
                )
            )  # fmt: skip
            rel.novos += 1
        fora = [e for e in sessao.scalars(select(EventoCorporativo)) if e.chave not in vistas]
        if substituir:
            for e in fora:
                sessao.delete(e)
            rel.removidos = len(fora)
        else:
            rel.fora_do_csv = len(fora)
    return rel
