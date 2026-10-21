"""Proventos do titular a partir da movimentação da B3 gravada no banco (engine/proventos).

A data do crédito de cada JCP (o ano em que ele é declarado) vem da correção do usuário
(``credito_jcp``) ou, sem ela, da data de aprovação no pacote público da B3
(``proventos_b3.csv``), casando empresa, data de pagamento e valor por ação. Sem nenhuma das
duas, o motor usa a alíquota retida para perceber o crédito no ano anterior (decisão do usuário de
20/10/2026)."""

from collections.abc import Iterable, Mapping
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from sqlalchemy import select

from coir import referencias
from coir.db.modelos import CreditoJcp, MovimentacaoB3
from coir.db.sessao import conectar
from coir.domain.b3 import LinhaMovimentacao
from coir.domain.parametros import ReferenciasFiscais
from coir.domain.proventos import LinhaB3, Provento
from coir.engine.proventos import JCP, proventos
from coir.erros import ImportacaoError

JCP_B3 = "JRS CAP PROPRIO"
TOLERANCIA_VALOR = Decimal("0.01")  # diferença relativa do valor por ação (arredondamento)


def linhas_banco(banco: Path, titular: int | None = None) -> list[LinhaB3]:
    consulta = select(MovimentacaoB3).order_by(MovimentacaoB3.data, MovimentacaoB3.id)
    if titular is not None:
        consulta = consulta.where(MovimentacaoB3.titular_id == titular)
    with conectar(banco) as fabrica, fabrica() as s:
        informados = {c.chave: c.data_credito for c in s.scalars(select(CreditoJcp))}
        return [
            LinhaB3(
                id=m.id,
                titular=m.titular_id,
                chave=m.chave,
                linha=LinhaMovimentacao(
                    data=m.data,
                    entrada=m.entrada,
                    movimentacao=m.movimentacao,
                    produto=m.produto,
                    codigo=m.codigo,
                    instituicao=m.instituicao,
                    quantidade=m.quantidade,
                    preco=m.preco,
                    valor=m.valor,
                ),
                credito=informados.get(m.chave),
                credito_origem="informada" if m.chave in informados else "",
            )
            for m in s.scalars(consulta)
        ]


def creditos_b3(
    linhas: Iterable[LinhaB3], publicos: Iterable[Mapping[str, str]],
    emissor_de: Mapping[str, str],
) -> list[LinhaB3]:  # fmt: skip
    """Completa a data do crédito dos JCP sem correção do usuário com a data de aprovação da B3:
    mesma empresa e data de pagamento; havendo mais de um (ON e PN), o de valor por ação mais
    próximo do preço da movimentação, dentro de 1%."""
    por_pagamento: dict[tuple[str, date], list[tuple[Decimal, date]]] = {}
    for r in publicos:
        if r["tipo"] == JCP_B3 and r["pagamento"]:
            chave = (r["emissor"], date.fromisoformat(r["pagamento"]))
            por_pagamento.setdefault(chave, []).append(
                (Decimal(r["valor"]), date.fromisoformat(r["aprovado_em"]))
            )
    saida: list[LinhaB3] = []
    for item in linhas:
        x = item.linha
        if item.credito is None and x.movimentacao == JCP and x.entrada:
            emissor = emissor_de.get(x.codigo, x.codigo[:4])
            candidatos = por_pagamento.get((emissor, x.data), [])
            if x.preco:
                preco = x.preco
                candidatos = sorted((c for c in candidatos
                                     if abs(c[0] - preco) <= TOLERANCIA_VALOR * preco),
                                    key=lambda c: abs(c[0] - preco))  # fmt: skip
            datas = {c[1] for c in candidatos[:1]} if x.preco else {c[1] for c in candidatos}
            if len(datas) == 1:
                item = item.model_copy(update={"credito": datas.pop(), "credito_origem": "B3"})
        saida.append(item)
    return saida


def proventos_banco(
    banco: Path, fiscais: ReferenciasFiscais, titular: int | None = None,
    pasta_refs: Path | None = None,
) -> tuple[list[Provento], list[str]]:  # fmt: skip
    avisos: list[str] = []
    pasta = pasta_refs or referencias.pasta_efetiva(None)
    emissor_de = {t: e for t, (e, _, _) in referencias.instrumentos(pasta).items()}
    publicos = referencias.proventos_b3(pasta)
    linhas = creditos_b3(linhas_banco(banco, titular), publicos, emissor_de)
    lista = proventos(linhas, fiscais.tipos, fiscais.parametros, avisos, por_pagamento(publicos))
    return lista, avisos


TIPO_B3 = {JCP_B3: "JCP", "DIVIDENDO": "DIVIDENDO"}


def por_pagamento(
    publicos: Iterable[Mapping[str, str]],
) -> dict[tuple[str, date], list[tuple[str, Decimal]]]:
    """Proventos do pacote da B3 por (emissor, pagamento): a origem dos reembolsos de BTC."""
    saida: dict[tuple[str, date], list[tuple[str, Decimal]]] = {}
    for r in publicos:
        tipo = TIPO_B3.get(r["tipo"])
        if tipo and r["pagamento"]:
            chave = (r["emissor"], date.fromisoformat(r["pagamento"]))
            item = (tipo, Decimal(r["valor"]))
            if item not in saida.setdefault(chave, []):
                saida[chave].append(item)
    return saida


def informar_credito(banco: Path, chave: str, data_credito: date | None, agora: datetime) -> None:
    """Grava (ou apaga, com None) a data do crédito de um JCP informada pelo usuário."""
    with conectar(banco) as fabrica, fabrica() as s, s.begin():
        linha = s.scalar(select(MovimentacaoB3).where(MovimentacaoB3.chave == chave))
        if linha is None:
            raise ImportacaoError("linha da movimentação não encontrada")
        if linha.movimentacao != JCP:
            raise ImportacaoError("só o JCP tem data de crédito")
        if data_credito is not None and data_credito > linha.data:
            raise ImportacaoError("o crédito não pode ser depois do pagamento")
        atual = s.scalar(select(CreditoJcp).where(CreditoJcp.chave == chave))
        if data_credito is None:
            if atual is not None:
                s.delete(atual)
        elif atual is None:
            s.add(CreditoJcp(titular_id=linha.titular_id, chave=chave,
                             data_credito=data_credito, registrado_em=agora))  # fmt: skip
        else:
            atual.data_credito, atual.registrado_em = data_credito, agora
