"""Proventos do titular a partir da movimentação da B3 gravada no banco (engine/proventos).

A data do crédito de cada JCP (o ano em que ele é declarado) vem da correção do usuário
(``credito_jcp``) ou, sem ela, do pacote público da B3: a mais tarde entre a aprovação e a data com
(``proventos_b3.csv``, casando empresa, data de pagamento e valor por ação; nos anos que ele não
alcança, ``proventos_historico_b3.csv``, casando empresa, classe, valor e data com). Sem nenhuma
das duas, o motor usa a alíquota retida para perceber o crédito no ano anterior (decisões do
usuário de 20 e 21/10/2026)."""

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
        creditos = select(CreditoJcp)
        if titular is not None:
            creditos = creditos.where(CreditoJcp.titular_id == titular)
        informados = {c.chave: c.data_credito for c in s.scalars(creditos)}
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
    emissor_de: Mapping[str, str], avisos: list[str] | None = None,
    historico: Iterable[Mapping[str, str]] = (), classe_de: Mapping[str, str] | None = None,
) -> list[LinhaB3]:  # fmt: skip
    """Completa a data do crédito dos JCP sem correção do usuário pela B3. O crédito é a mais tarde
    entre a aprovação e a data com: antes da data com o acionista não está definido e não há
    crédito individualizado (IN SRF 41/1998 art. 3º; IN RFB 1.700/2017 art. 75 §10; Lei
    6.404/1976 art. 205; decisão do usuário de 21/10/2026). Casamento: mesma empresa e data de
    pagamento no pacote; havendo mais de um (ON e PN), o de valor por ação mais próximo do preço
    da movimentação, dentro de 1%. Sem pagamento no pacote (anos que ele não alcança), vale o
    ``historico`` (sem data de pagamento): mesma empresa e classe, valor por ação dentro de 1% e a
    data com mais recente antes do pagamento; empate com créditos diferentes fica sem crédito.
    Crédito depois do pagamento (dado público retificado) é ignorado, com aviso; a data informada
    pelo usuário antes da data com da B3 vale, com aviso (o informe da companhia prevalece)."""
    por_pagamento: dict[tuple[str, date], list[tuple[Decimal, date, date | None]]] = {}
    for r in publicos:
        if r["tipo"] == JCP_B3 and r["pagamento"]:
            pagamento, aprovado = date.fromisoformat(r["pagamento"]), date.fromisoformat(
                r["aprovado_em"])  # fmt: skip
            data_com = date.fromisoformat(r["data_com"]) if r.get("data_com") else None
            credito = max(aprovado, data_com or aprovado)
            if credito > pagamento:
                if avisos is not None:
                    avisos.append(f"pacote da B3: JCP de {r['emissor']} pago em {pagamento} com "
                                  f"crédito em {credito}, ignorado")  # fmt: skip
                continue
            por_pagamento.setdefault((r["emissor"], pagamento), []).append(
                (Decimal(r["valor"]), credito, data_com)
            )
    por_classe: dict[tuple[str, str], list[tuple[date, Decimal, date]]] = {}
    for r in historico:
        if r["tipo"] == JCP_B3:
            com = date.fromisoformat(r["data_com"])
            por_classe.setdefault((r["emissor"], r["classe"]), []).append(
                (com, Decimal(r["valor"]), max(com, date.fromisoformat(r["aprovado_em"])))
            )
    saida: list[LinhaB3] = []
    for item in linhas:
        x = item.linha
        if x.movimentacao == JCP and x.entrada:
            emissor = emissor_de.get(x.codigo, x.codigo[:4])
            classe = (classe_de or {}).get(x.codigo, "")
            achado = _casar(x.data, x.preco, por_pagamento.get((emissor, x.data), []),
                            por_classe.get((emissor, classe), []))  # fmt: skip
            if achado is not None and item.credito is None:
                credito, _, origem = achado
                item = item.model_copy(update={"credito": credito, "credito_origem": origem})
            elif achado is not None and achado[1] and item.credito and item.credito < achado[1]:
                if avisos is not None:
                    avisos.append(f"JCP de {x.codigo} pago em {x.data:%d/%m/%Y}: crédito informado "
                                  f"em {item.credito:%d/%m/%Y}, antes da data com "
                                  f"{achado[1]:%d/%m/%Y} da B3 (confira o informe)")  # fmt: skip
        saida.append(item)
    return saida


def _casar(
    pagamento: date, preco: Decimal | None, do_pacote: list[tuple[Decimal, date, date | None]],
    do_historico: list[tuple[date, Decimal, date]],
) -> tuple[date, date | None, str] | None:  # fmt: skip
    """(crédito, data com, origem) do JCP pago em ``pagamento``: o pacote (pela data de
    pagamento) e, sem ele, o histórico; None sem casamento único."""
    candidatos = do_pacote
    if preco:
        candidatos = sorted((c for c in do_pacote if abs(c[0] - preco) <= TOLERANCIA_VALOR * preco),
                            key=lambda c: abs(c[0] - preco))[:1]  # fmt: skip
    if len({(c[1], c[2]) for c in candidatos}) == 1:
        return candidatos[0][1], candidatos[0][2], "B3"
    if candidatos or not preco:
        return None
    tolerancia = TOLERANCIA_VALOR * preco
    linhas = [(com, credito) for com, valor, credito in do_historico
              if com < pagamento and abs(valor - preco) <= tolerancia]  # fmt: skip
    if not linhas:
        return None
    # o JCP mensal repete o valor: o deste pagamento é o de data com mais recente antes dele
    ultima = max(com for com, _ in linhas)
    creditos = {credito for com, credito in linhas if com == ultima}
    if len(creditos) != 1:
        return None
    credito = creditos.pop()
    return (credito, ultima, "B3, histórico") if credito <= pagamento else None


def proventos_banco(
    banco: Path, fiscais: ReferenciasFiscais, titular: int | None = None,
    pasta_refs: Path | None = None,
) -> tuple[list[Provento], list[str]]:  # fmt: skip
    avisos: list[str] = []
    pasta = pasta_refs or referencias.pasta_efetiva(None)
    cadastro = referencias.instrumentos(pasta)
    emissor_de = {t: e for t, (e, _, _) in cadastro.items()}
    classe_de = {t: (s.split() or [""])[0] for t, (_, s, _) in cadastro.items()}
    publicos = referencias.proventos_b3(pasta)
    linhas = creditos_b3(linhas_banco(banco, titular), publicos, emissor_de, avisos,
                         referencias.proventos_historico_b3(pasta), classe_de)  # fmt: skip
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
    informar_creditos(banco, [(chave, data_credito)], agora)


def informar_creditos(
    banco: Path, datas: Iterable[tuple[str, date | None]], agora: datetime
) -> None:
    """Várias datas de crédito numa transação só (as propostas pela DIRPF): erro em uma, nada é
    gravado."""
    with conectar(banco) as fabrica, fabrica() as s, s.begin():
        for chave, data_credito in datas:
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
