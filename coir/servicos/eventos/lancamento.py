"""Gravar e excluir eventos (um a um ou em lote), sempre conferidos no motor, e o lançamento
automático dos eventos seguros."""

import re
from collections.abc import Sequence
from dataclasses import dataclass, replace
from datetime import date, datetime
from pathlib import Path

from sqlalchemy import select
from sqlalchemy.orm import Session

from coir import auditoria
from coir.db import repositorio as repo
from coir.db.modelos import (
    Conta,
    EventoCorporativo,
    PropostaIgnorada,
)
from coir.db.sessao import conectar
from coir.domain.fiscal import (
    NaturezaEvento,
    TipoEvento,
    Visao,
)
from coir.domain.parametros import ReferenciasFiscais
from coir.domain.referencias import Referencias
from coir.engine.movimentos import movimentar
from coir.erros import CoirError, ImportacaoError, MotorError, ReferenciaError
from coir.servicos import trilha
from coir.servicos.apuracao import entradas
from coir.servicos.eventos.dados import (
    DadosEvento,
    _ativos_do_ticker,
    _campos_chave,
    _chaves,
    _preencher,
    _validar,
)
from coir.servicos.eventos.sugestoes import propostas


def salvar(
    banco: Path, titular: int, d: DadosEvento, refs: Referencias, fiscais: ReferenciasFiscais,
    evento_id: int | None = None,
) -> int:  # fmt: skip
    """Inclui (ou altera) o evento numa transação e confere com o motor (visão por titular e por
    conta): evento que não fecha com as operações (sem posição, quantidade maior que a posição...)
    não é gravado e o erro diz por quê. Devolve o id."""
    [novo] = salvar_lote(banco, titular, [(d, evento_id)], refs, fiscais)
    return novo


def salvar_lote(
    banco: Path, titular: int, lote: list[tuple[DadosEvento, int | None]], refs: Referencias,
    fiscais: ReferenciasFiscais,
) -> list[int]:  # fmt: skip
    """Vários eventos (novos ou alterados) numa transação, conferidos juntos pelo motor no fim:
    a cisão de uma posição em várias contas só fecha com os eventos de todas."""
    feitos: list[int] = []
    with conectar(banco) as fabrica, fabrica.begin() as s:
        antes = pendencia_antes(s, titular, refs, fiscais)
        for d, evento_id in lote:
            feitos.append(_gravar(s, titular, d, evento_id, refs))
        conferir_no_motor(s, titular, refs, fiscais, antes)
    for d, evento_id in lote:
        trilha.evento_salvo(titular, d, evento_id is not None)
    return feitos


# o direito recebido e exercido pode nunca ter passado por uma nota de corretagem
SEM_NEGOCIO_NO_BANCO = frozenset({
    TipoEvento.CUSTODIA_INICIAL, TipoEvento.OFERTA_PUBLICA, TipoEvento.DIREITO_RECEBIDO,
    TipoEvento.EXERCICIO_DIREITO,
})  # fmt: skip


RE_TICKER_CUSTODIA = re.compile(r"^[A-Z0-9]{4}\d{1,2}$")  # ação, unit, ETF, BDR, FII


def _gravar(
    s: Session, titular: int, d: DadosEvento, evento_id: int | None, refs: Referencias
) -> int:
    d = replace(d, ticker=d.ticker.strip().upper(),
                ticker_destino=(d.ticker_destino or "").strip().upper() or None,
                observacao=d.observacao.strip())  # fmt: skip
    _validar(d)
    conta = s.get(Conta, d.conta)
    destino = None if d.conta_destino is None else s.get(Conta, d.conta_destino)
    if conta is None or conta.titular_id != titular:
        raise ImportacaoError("conta de outro titular ou inexistente")
    if d.conta_destino is not None and (destino is None or destino.titular_id != titular):
        raise ImportacaoError("transferência entre titulares diferentes")
    chaves = _chaves(_campos_chave(conta, destino, d), 0)
    repetido = s.scalar(select(EventoCorporativo).where(EventoCorporativo.chave.in_(chaves)))
    if repetido is not None and repetido.id != evento_id:
        raise ImportacaoError(f"evento igual já lançado (id {repetido.id})")
    registro = EventoCorporativo() if evento_id is None else s.get(EventoCorporativo, evento_id)
    if registro is None or (evento_id is not None and registro.conta.titular_id != titular):
        raise ImportacaoError(f"evento {evento_id} não encontrado")
    # código da época (CNTO3 em 2019): é o mesmo ativo do código novo (SBFG3), que é como as notas
    # podem ter entrado; sem isso, a custódia abriria um ativo separado das operações
    atual = refs.ticker_na_data(d.ticker, d.data)
    ativo = (
        (repo.ativo_existente(s, atual, d.data) if atual != d.ticker else None)
        or repo.ativo_existente(s, d.ticker, d.data)
        or repo.ativo_pelo_codigo_antigo(s, d.ticker, d.data)
    )
    outra_classe = d.origem == "b3" and d.natureza is NaturezaEvento.OUTRA_CLASSE
    if ativo is None and (outra_classe or d.tipo in SEM_NEGOCIO_NO_BANCO):
        # quem começa sem as notas antigas, compra na oferta, recebe outra classe da companhia
        # ou recebe e exerce um direito de subscrição lança ativo ainda sem negócio no banco
        if not RE_TICKER_CUSTODIA.match(d.ticker):
            raise ImportacaoError(f"{d.ticker}: código de negociação inválido (ex.: PETR4)")
        ativo = repo.Ativos(s).ativo_id(atual, d.data)
    if ativo is None:
        raise ImportacaoError(f"{d.ticker} sem negócio no banco em {d.data:%d/%m/%Y}")
    _preencher(s, registro, d, chaves[0], ativo)
    if evento_id is None:
        s.add(registro)
    s.flush()
    return registro.id


def _erro_do_livro(
    s: Session, titular: int, refs: Referencias, fiscais: ReferenciasFiscais
) -> MotorError | None:  # fmt: skip
    """Onde o motor para com o banco como está na sessão (None: o livro fecha): a pendência mais
    antiga das duas visões. Parar na de titular não pode esconder a de conta (evento lançado na
    conta errada, simulação da 3.1.0 com as notas reais)."""
    notas, eventos = entradas(s, refs, fiscais)
    notas = [n for n in notas if n.titular == titular]
    eventos = [e for e in eventos if e.titular == titular]
    erros: list[MotorError] = []
    for visao in (Visao.TITULAR, Visao.CONTA):
        try:
            movimentar(notas, eventos, fiscais.parametros, date.max, visao)
        except MotorError as erro:
            erros.append(erro)
    return min(erros, key=lambda e: e.data or date.min, default=None)


def pendencia_antes(
    s: Session, titular: int, refs: Referencias, fiscais: ReferenciasFiscais
) -> MotorError | None:  # fmt: skip
    try:
        return _erro_do_livro(s, titular, refs, fiscais)
    except ReferenciaError:
        return None  # o motor recusa de novo depois da mudança, com a mensagem


def conferir_no_motor(
    s: Session, titular: int, refs: Referencias, fiscais: ReferenciasFiscais,
    antes: MotorError | None, o_que: str = "o evento",
) -> None:  # fmt: skip
    """A mudança não pode criar pendência no livro. Com uma pendência de antes (ex.: bonificação
    que falta noutro ativo), o livro pode continuar parando nela ou passar dela e parar numa data
    posterior (a mudança resolveu aquela, a próxima é outra); assim o usuário corrige uma por vez,
    sem a interface travar na primeira."""
    try:
        depois = _erro_do_livro(s, titular, refs, fiscais)
    except ReferenciaError as erro:
        raise ImportacaoError(f"{o_que} não fecha com as operações: {erro}") from erro
    if depois is None or (
        antes is not None
        and (
            str(depois) == str(antes)
            or (antes.data is not None and depois.data is not None and depois.data > antes.data)
        )
    ):
        return  # fmt: skip
    raise ImportacaoError(f"{o_que} não fecha com as operações: {depois}") from depois


def excluir(
    banco: Path, titular: int, evento_id: int, refs: Referencias, fiscais: ReferenciasFiscais,
    pasta: Path | None = None,
) -> None:  # fmt: skip
    """Apaga o evento se o livro continuar fechando sem ele (ex.: venda que dependia de um
    desdobramento impede apagar o desdobramento)."""
    excluir_lote(banco, titular, [evento_id], refs, fiscais, pasta=pasta)


def excluir_lote(
    banco: Path, titular: int, ids: Sequence[int], refs: Referencias, fiscais: ReferenciasFiscais,
    ignorar_ausentes: bool = False, pasta: Path | None = None,
) -> None:  # fmt: skip
    """Apaga os eventos numa transação, conferidos juntos pelo motor no fim (desfaz um lote; um
    evento do lote pode já ter sido excluído na tela: ``ignorar_ausentes``). Com ``pasta``, a
    proposta do evento apagado fica ignorada: o lançamento automático não o traz de volta."""
    apagados: list[tuple[int, str, date]] = []
    with conectar(banco) as fabrica, fabrica.begin() as s:
        antes = pendencia_antes(s, titular, refs, fiscais)
        registros: list[EventoCorporativo] = []
        for evento_id in ids:
            registro = s.get(EventoCorporativo, evento_id)
            if registro is None and ignorar_ausentes:
                continue
            if registro is None or registro.conta.titular_id != titular:
                raise ImportacaoError(f"evento {evento_id} não encontrado")
            registros.append(registro)
        for registro in registros:
            s.delete(registro)
        s.flush()
        try:
            conferir_no_motor(s, titular, refs, fiscais, antes)
        except ImportacaoError as erro:
            raise ImportacaoError(f"sem este evento o livro não fecha: {erro.__cause__}") from erro
        for registro in registros:
            auditoria.evento("eventos.excluido", titular=titular, tipo=str(registro.tipo),
                             data=registro.data, origem=registro.origem)  # fmt: skip
            if registro.origem == "b3":
                apagados.append((registro.ativo_id, str(registro.tipo), registro.data))
    if pasta is not None and apagados:
        _ignorar_as_apagadas(banco, titular, apagados, refs, fiscais, pasta)


def _ignorar_as_apagadas(
    banco: Path, titular: int, apagados: list[tuple[int, str, date]], refs: Referencias,
    fiscais: ReferenciasFiscais, pasta: Path,
) -> None:  # fmt: skip
    lista = propostas(banco, titular, refs, fiscais, pasta)
    with conectar(banco) as fabrica, fabrica.begin() as s:
        for ativo, tipo, dia in apagados:
            for p in lista:
                if (p.situacao == "pendente" and p.data == dia and str(p.dados.tipo) == tipo
                        and ativo in _ativos_do_ticker(s, p.dados.ticker)):  # fmt: skip
                    s.add(PropostaIgnorada(titular_id=titular, chave=p.chave,
                                           ignorada_em=datetime.now()))  # fmt: skip


@dataclass(frozen=True)
class ResultadoLote:
    ids: tuple[int, ...]
    recusadas: tuple[tuple[str, str], ...]  # (ativo, motivo)


def lancar_seguras(
    banco: Path, titular: int, chaves: Sequence[str], refs: Referencias,
    fiscais: ReferenciasFiscais, pasta: Path,
) -> ResultadoLote:  # fmt: skip
    """Lança de uma vez as propostas escolhidas que continuam seguras e pendentes (a tela pode estar
    atrás do banco). A que o motor recusa volta com o motivo e não trava as outras."""
    faltam = set(chaves)
    lista = [p for p in propostas(banco, titular, refs, fiscais, pasta) if p.chave in faltam]
    lote = [p for p in lista if p.segura and p.situacao == "pendente"]
    recusadas = [(p.codigo, "agora precisa de conferência: lance na lista") for p in lista
                 if p not in lote]  # fmt: skip
    ids: list[int] = []
    try:
        pedidos: list[tuple[DadosEvento, int | None]] = [(p.dados, None) for p in lote]
        ids.extend(salvar_lote(banco, titular, pedidos, refs, fiscais))
    except CoirError:  # uma não fecha: as outras entram uma a uma, e ela sai com o motivo
        for p in lote:
            try:
                ids.append(salvar(banco, titular, p.dados, refs, fiscais))
            except CoirError as erro:
                recusadas.append((p.codigo, str(erro)))
    return ResultadoLote(tuple(ids), tuple(recusadas))


def lancar_automaticos(
    banco: Path, titular: int, refs: Referencias, fiscais: ReferenciasFiscais, pasta: Path
) -> ResultadoLote:
    """Lança sozinho o que não pede nada ao usuário: as propostas seguras de ativos em carteira até
    a última nota dele. Em rodadas, porque só a mais antiga de cada ativo é segura (a seguinte
    depende da quantidade depois dela). O que o usuário apagou fica ignorado e não volta."""
    ids: list[int] = []
    recusadas: list[tuple[str, str]] = []
    tentadas: set[str] = set()
    while True:
        lista = [p.chave for p in propostas(banco, titular, refs, fiscais, pasta)
                 if p.segura and p.situacao == "pendente" and p.chave not in tentadas]  # fmt: skip
        if not lista:
            return ResultadoLote(tuple(ids), tuple(recusadas))
        tentadas.update(lista)
        rodada = lancar_seguras(banco, titular, lista, refs, fiscais, pasta)
        ids.extend(rodada.ids)
        recusadas.extend(rodada.recusadas)
