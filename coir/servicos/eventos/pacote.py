"""Propostas tiradas do pacote público de referências: eventos da bolsa e da CVM que atingem a
posição na data com, ajuste das séries de opção e restituições de capital."""

import hashlib
from collections import defaultdict
from dataclasses import replace
from datetime import date, timedelta
from decimal import ROUND_FLOOR, ROUND_HALF_UP, Decimal
from pathlib import Path

from sqlalchemy import select
from sqlalchemy.orm import Session

from coir import fontes
from coir.db.modelos import (
    Conta,
    SerieOpcao,
    TickerHistorico,
)
from coir.domain.fiscal import (
    AtivoIsolado,
    Movimento,
    NaturezaEvento,
    TipoAtivo,
    TipoEvento,
    Visao,
)
from coir.domain.formato import reais
from coir.domain.parametros import ReferenciasFiscais
from coir.domain.referencias import Referencias
from coir.engine.calendario import proximo_pregao
from coir.engine.movimentos import movimentar
from coir.erros import MotorError, ReferenciaError
from coir.servicos.apuracao import Entradas, entradas
from coir.servicos.eventos.dados import DadosEvento, Proposta, _ja_lancado

# eventos do pacote público (rótulos como a B3 escreve) -> tipo do evento
TIPOS_PACOTE = {
    "DESDOBRAMENTO": TipoEvento.DESDOBRAMENTO,
    "GRUPAMENTO": TipoEvento.GRUPAMENTO,
    "BONIFICACAO": TipoEvento.BONIFICACAO,
    "SUBSCRICAO": TipoEvento.SUBSCRICAO,
    # dinheiro junto de uma incorporação: "fator" é o % do custo que sai das ações recebidas,
    # "preco" é o dinheiro por ação e "data_subscricao" é o pagamento (PETZ3 -> AUAU3, resgate)
    "REDUCAO_CAPITAL": TipoEvento.REDUCAO_CAPITAL,
}


AJUSTA_SERIE = {t: t for t in (TipoEvento.DESDOBRAMENTO, TipoEvento.GRUPAMENTO,
                                TipoEvento.BONIFICACAO)}  # fmt: skip


GRUPAMENTO_EXTREMO = Decimal("0.01")  # 100 viram 1 ou mais forte


TIPOS_SEGUROS = frozenset({TipoEvento.DESDOBRAMENTO, TipoEvento.GRUPAMENTO})


FONTE_CUSTO_SEGURA = "cvm_ipe"  # origem do evento: comunicado da companhia na CVM


AVISO_GRUPAMENTO_EXTREMO = (
    "confira o aviso aos acionistas: grupamento tão forte costuma vir seguido de desdobramento "
    "no mesmo dia (ex.: Itaú em 2011), sem mudar a quantidade; a B3 publica só o grupamento"
)


NOTA_PACOTE = "evento do pacote público da B3: data ex = pregão seguinte à data com {data_com}"


class _HistoricoTickers:
    """O histórico de tickers lido uma vez (as propostas do pacote consultavam a tabela a cada
    evento): ticker -> ativos e ativo -> tickers."""

    def __init__(self, s: Session) -> None:
        self._ativos: dict[str, set[int]] = defaultdict(set)
        self._tickers: dict[int, set[str]] = defaultdict(set)
        for ticker, ativo in s.execute(select(TickerHistorico.ticker, TickerHistorico.ativo_id)):
            self._ativos[ticker].add(ativo)
            self._tickers[ativo].add(ticker)

    def ativos(self, ticker: str) -> set[int]:
        return set(self._ativos.get(ticker, ()))

    def nomes(self, ticker: str, ativos: set[int]) -> set[str]:
        """O ticker e todos os nomes que os ativos dele já tiveram."""
        return {ticker} | {t for a in ativos for t in self._tickers.get(a, ())}


def _posicoes(movs: list[Movimento]) -> dict[tuple[int, int], list[tuple[date, Decimal]]]:
    """(conta, ativo) -> [(data, quantidade depois)], na visão por conta, em ordem."""
    saida: dict[tuple[int, int], list[tuple[date, Decimal]]] = defaultdict(list)
    for m in movs:
        if m.conta is not None and m.ativo is not None:
            saida[(m.conta, m.ativo)].append((m.data, m.quantidade_apos))
    return saida


def _custos(movs: list[Movimento]) -> dict[tuple[int, int], list[tuple[date, Decimal]]]:
    """(conta, ativo) -> [(data, custo total depois)], na visão por conta, em ordem."""
    saida: dict[tuple[int, int], list[tuple[date, Decimal]]] = defaultdict(list)
    for m in movs:
        if m.conta is not None and m.ativo is not None:
            saida[(m.conta, m.ativo)].append((m.data, m.custo_apos))
    return saida


def _na_data(historico: list[tuple[date, Decimal]], dia: date) -> Decimal:
    q = Decimal(0)
    for quando, depois in historico:
        if quando > dia:
            break
        q = depois
    return q


def _do_pacote(
    s: Session, titular: int, refs: Referencias, fiscais: ReferenciasFiscais, pasta: Path,
    contas: list[Conta], ignoradas: set[str], lancados: list[tuple[int, int | None, date]],
    base: Entradas | None = None,
) -> list[Proposta]:  # fmt: skip
    """Para cada evento do pacote de ativo que o titular tinha na data com (por conta), o evento
    com a quantidade pelo fator e, nas séries de opção abertas do ativo, o ajuste de quantidade."""
    eventos = fontes.eventos(pasta)
    if not eventos:
        return []
    por_isin = fontes.tickers_por_isin(pasta)
    notas, evs = base if base is not None else entradas(s, refs, fiscais)
    notas = [n for n in notas if n.titular == titular]
    evs = [e for e in evs if e.titular == titular]
    limite = date.max
    isolados: list[AtivoIsolado] = []
    try:
        # o ativo que não fecha sai do cálculo a partir do problema, e os outros seguem: um ativo
        # travado não pode esconder os eventos de toda a carteira
        movs = movimentar(notas, evs, fiscais.parametros, limite, Visao.CONTA, isolados=isolados)
    except MotorError as erro:
        if erro.data is None:
            raise
        # o livro parou (ex.: exercício com a quantidade já ajustada por uma bonificação que falta):
        # as posições até a véspera bastam para propor o evento que destrava o livro
        limite = erro.data
        movs = movimentar(notas, evs, fiscais.parametros, limite - timedelta(days=1), Visao.CONTA)
    parou_em = {i.ativo: i.desde for i in isolados}
    historico = _posicoes(movs)
    custos = _custos(movs)
    inicio = min((m.data for m in movs), default=date.max)
    tickers_serie = {m.ativo: m.ticker for m in movs if m.tipo is TipoAtivo.OPCAO and m.ativo}
    objeto_da_serie = {
        s_.serie: s_.serie[:4] + refs.classes[s_.classe]
        for s_ in s.scalars(select(SerieOpcao))
        if s_.classe in refs.classes
    }
    nome_conta = {c.id: c.corretora.nome for c in contas}
    historico_tickers = _HistoricoTickers(s)
    saida: list[Proposta] = []
    for ev in eventos:
        tipo = TIPOS_PACOTE.get(ev.tipo)
        cisao = ev.tipo.startswith("CIS") and ev.isin_emitido != ev.isin
        ticker = por_isin.get(ev.isin)
        if ticker is None or (tipo is None and not cisao):
            continue
        data_com = date.fromisoformat(ev.data_com)  # registros já validados (fontes.eventos)
        cal = fiscais.calendario
        if not cal.primeiro_ano <= data_com.year <= cal.ultimo_ano or data_com < inicio:
            continue  # antes do primeiro negócio (ou fora do calendário): não atinge a carteira
        try:
            data_ex = proximo_pregao(data_com, cal)
        except ReferenciaError:
            continue  # data ex depois do fim do calendário de feriados
        if data_ex > limite:
            continue  # depois de onde o livro parou: a posição na data com não é conhecida
        # a B3 escreve o grupamento como razão (0,1 = 10 viram 1) e os demais em % (200 = +200%)
        fator = Decimal(ev.fator) / (1 if tipo is TipoEvento.GRUPAMENTO else 100)
        ativos = historico_tickers.ativos(ticker)
        nomes = historico_tickers.nomes(ticker, ativos)
        ajusta = tipo is not None and tipo in AJUSTA_SERIE
        for (conta, ativo), hist in historico.items():
            serie = tickers_serie.get(ativo)
            # só as posições que o evento atinge (a ação dele ou a série de opção sobre ela): a
            # posição na data com, mais cara, sai só delas
            if serie is None:
                atingida = ativo in ativos
            else:
                objeto = objeto_da_serie.get(serie) or refs.opcoes_raiz.get(serie[:4])
                atingida = ajusta and objeto in nomes
            if not atingida:
                continue
            q = _na_data(hist, data_com)
            if not q or conta not in nome_conta or data_ex > parou_em.get(ativo, date.max):
                continue  # depois de onde o ativo parou, a posição na data com não é conhecida
            if ativo in ativos and serie is None and tipo is TipoEvento.REDUCAO_CAPITAL:
                dados, nota = _resgate(ev, ticker, conta, q, _na_data(custos[(conta, ativo)],
                                                                      data_com), fator, data_ex,
                                       por_isin)  # fmt: skip
            elif ativo in ativos and serie is None:
                dados, nota = _evento_acao(ev, tipo, cisao, ticker, conta, q, fator, data_ex,
                                           por_isin)  # fmt: skip
            elif (serie is not None
                  and (objeto_da_serie.get(serie) or refs.opcoes_raiz.get(serie[:4])) in nomes
                  and tipo is not None and tipo in AJUSTA_SERIE):  # fmt: skip
                dados, nota = _ajuste_serie(AJUSTA_SERIE[tipo], serie, conta, q, fator, data_ex)
            else:
                continue
            if dados is None:
                continue
            chave = hashlib.sha256(f"b3pub|{ev.isin}|{ev.tipo}|{ev.data_com}|"
                                   f"{ev.isin_emitido}|{ev.fator}|{conta}|{dados.ticker}"
                                   .encode()).hexdigest()  # fmt: skip
            if chave in ignoradas:
                situacao = "ignorada"
            elif _ja_lancado(s, refs, dados.ticker, dados.data, lancados,
                             troca_conta=tipo is not TipoEvento.REDUCAO_CAPITAL):  # fmt: skip
                situacao = "lançado"
            else:
                situacao = "pendente"
            saida.append(Proposta(
                chave, dados.data, f"{ev.tipo.capitalize()} (eventos da B3)", dados.ticker, q,
                None, nome_conta[conta], situacao, dados,
                "; ".join(x for x in (NOTA_PACOTE.format(data_com=f"{data_com:%d/%m/%Y}"), nota)
                          if x),
                segura=(serie is None and ev.isin_emitido in ("", ev.isin)
                        and _segura(dados, fator, ev.origem)),
                custo_por_acao=(Decimal(ev.custo_atribuido) if serie is None
                                and dados.tipo is TipoEvento.BONIFICACAO
                                and ev.custo_atribuido else None),
                referencia=ev.referencia))  # fmt: skip
    saida += _propor_restituicoes(s, refs, fiscais, pasta, por_isin, historico, nome_conta,
                                  tickers_serie, parou_em, inicio, limite, ignoradas,
                                  lancados, historico_tickers)  # fmt: skip
    return saida


REST_CAP = "REST CAP DIN"  # restituição de capital em dinheiro nos proventos da B3


NOTA_RESTITUICAO = ("restituição de capital em dinheiro do pacote público da B3 (R$ {valor} por "
                    "ação, data com {data_com}): reduz o custo, sem resultado (IN RFB 1.585/2015 "
                    "art. 58 §8º); confira o valor creditado na movimentação da B3")  # fmt: skip


def _restituicoes(pasta: Path, refs: Referencias,
                  por_isin: dict[str, str]) -> list[tuple[str, date, Decimal]]:  # fmt: skip
    """(ticker, data com, valor por ação) das restituições de capital em dinheiro do pacote: as
    recentes (com ISIN) e as do histórico (empresa e classe); a mesma nas duas entra uma vez."""
    achadas: dict[tuple[str, str], Decimal] = {}
    for r in fontes.proventos(pasta):
        ticker = por_isin.get(r.isin)
        if r.tipo == REST_CAP and ticker and r.data_com:
            achadas.setdefault((ticker, r.data_com), Decimal(r.valor))
    for r in fontes.proventos_historicos(pasta):
        digito = refs.classes.get(r.classe)
        if r.tipo == REST_CAP and digito:
            achadas.setdefault((r.emissor + digito, r.data_com), Decimal(r.valor))
    saida: list[tuple[str, date, Decimal]] = []
    for (ticker, data_com), valor in sorted(achadas.items()):
        try:
            saida.append((ticker, date.fromisoformat(data_com), valor))
        except ValueError:
            continue
    return saida


def _propor_restituicoes(
    s: Session, refs: Referencias, fiscais: ReferenciasFiscais, pasta: Path,
    por_isin: dict[str, str], historico: dict[tuple[int, int], list[tuple[date, Decimal]]],
    nome_conta: dict[int, str], tickers_serie: dict[int, str], parou_em: dict[int, date],
    inicio: date, limite: date, ignoradas: set[str], lancados: list[tuple[int, int | None, date]],
    historico_tickers: _HistoricoTickers,
) -> list[Proposta]:  # fmt: skip
    """Restituição de capital em dinheiro como redução de capital só com o valor (posição na data
    com vezes o valor por ação, truncado no centavo), na data ex. Pede conferência: a B3 credita
    o valor da conta, que a movimentação mostra."""
    cal = fiscais.calendario
    saida: list[Proposta] = []
    for ticker, data_com, valor in _restituicoes(pasta, refs, por_isin):
        if not cal.primeiro_ano <= data_com.year <= cal.ultimo_ano or data_com < inicio:
            continue
        try:
            data_ex = proximo_pregao(data_com, cal)
        except ReferenciaError:
            continue
        if data_ex > limite:
            continue
        ativos = historico_tickers.ativos(ticker)
        for (conta, ativo), hist in historico.items():
            if ativo not in ativos or ativo in tickers_serie:
                continue  # posição de outro ativo: nem calcula a da data com
            q = _na_data(hist, data_com)
            if q <= 0 or conta not in nome_conta or data_ex > parou_em.get(ativo, date.max):
                continue  # fmt: skip
            total = (q * valor).quantize(Decimal("0.01"), rounding=ROUND_FLOOR)
            if total <= 0:
                continue
            dados = DadosEvento(
                tipo=TipoEvento.REDUCAO_CAPITAL, data=data_ex, conta=conta, ticker=ticker, cv="V",
                quantidade=None, valor=total, origem="b3",
                observacao="pacote público da B3: restituição de capital em dinheiro",
            )  # fmt: skip
            chave = hashlib.sha256(f"b3rest|{ticker}|{data_com}|{valor}|{conta}"
                                   .encode()).hexdigest()  # fmt: skip
            if chave in ignoradas:
                situacao = "ignorada"
            elif _ja_lancado(s, refs, ticker, data_ex, lancados, troca_conta=False):
                situacao = "lançado"
            else:
                situacao = "pendente"
            saida.append(Proposta(
                chave, data_ex, "Restituição de capital (proventos da B3)", ticker, q, total,
                nome_conta[conta], situacao, dados,
                NOTA_RESTITUICAO.format(valor=valor, data_com=f"{data_com:%d/%m/%Y}")))  # fmt: skip
    return saida


def _segura(d: DadosEvento, fator: Decimal, origem: str) -> bool:
    """Proposta do pacote público (já com a data ex) que não pede nada ao usuário: desdobramento e
    grupamento de ação, na conta certa, com quantidade, e bonificação com o custo por ação lido do
    documento da companhia na CVM. Fica de fora o grupamento extremo, que costuma vir seguido de
    desdobramento no mesmo dia (ex.: Itaú em 2011)."""
    if d.conta is None or not d.quantidade:
        return False
    if d.tipo is TipoEvento.BONIFICACAO:
        return bool(d.valor) and FONTE_CUSTO_SEGURA in origem
    if d.tipo not in TIPOS_SEGUROS:
        return False
    return d.tipo is not TipoEvento.GRUPAMENTO or fator > GRUPAMENTO_EXTREMO


def _fator(ev: fontes.Evento, tipo: TipoEvento | None) -> str:
    """Como a B3 escreve: grupamento em razão (0,1 = 10 viram 1); os demais em %."""
    return f"razão {ev.fator}" if tipo is TipoEvento.GRUPAMENTO else f"{ev.fator}%"


def _fracao(total: Decimal) -> tuple[Decimal, str]:
    """Quantidade do evento com a fração: ela fica na posição até o leilão, que a vende pelo custo
    médio (IN RFB 1.585/2015 art. 58 caput; Lei 6.404/1976 art. 169 §3º; critério do COIR de
    10/11/2026). Sem isso, o leilão tirava a fração das ações inteiras e a posição ficava
    fracionada."""
    inteiro = total.to_integral_value(ROUND_FLOOR)
    nota = (f"fração de {total - inteiro} fica na posição até o leilão: lance o leilão de fração "
            "quando for creditado") if total != inteiro else ""  # fmt: skip
    return total, nota


def _resgate(
    ev: fontes.Evento, ticker: str, conta: int, q: Decimal, custo: Decimal, parte: Decimal,
    data_ex: date, por_isin: dict[str, str],
) -> tuple[DadosEvento | None, str]:  # fmt: skip
    """Dinheiro junto de uma incorporação: a parte do custo que era da ação resgatada sai das ações
    recebidas, sem resultado na bolsa (SC Cosit 37/2020; parecer do especialista de 07/11/2026). O
    valor do evento é esse custo, nunca o dinheiro. O resgate é alienação fora de bolsa."""
    if q <= 0:
        return None, ""
    destino = por_isin.get(ev.isin_emitido, ticker)
    valor = (custo * parte).quantize(Decimal("0.01"), ROUND_HALF_UP)
    dinheiro = (q * Decimal(ev.preco or 0)).quantize(Decimal("0.01"), ROUND_HALF_UP)
    pago = date.fromisoformat(ev.data_subscricao) if ev.data_subscricao else data_ex
    nota = (
        f"o valor proposto é a parte do custo que sai de {destino} ({ev.fator}%), não o "
        f"dinheiro. O dinheiro ({reais(dinheiro)}, pago em {pago:%d/%m/%Y}) é resgate de "
        "ações: menos este custo, dá "
        "ganho ou perda fora da bolsa, no programa GCAP da Receita. O COIR não calcula esse "
        "resultado, a perda não compensa ganhos da bolsa, e há isenção até R$ 35 mil no mês"
    )
    return DadosEvento(
        tipo=TipoEvento.REDUCAO_CAPITAL, data=data_ex, conta=conta, ticker=destino, cv="V",
        quantidade=None, valor=valor, origem="b3",
        observacao=f"pacote público da B3: parte do custo na incorporação ({ev.fator}%)",
    ), nota  # fmt: skip


def _evento_acao(
    ev: fontes.Evento, tipo: TipoEvento | None, cisao: bool, ticker: str, conta: int,
    q: Decimal, fator: Decimal, data_ex: date, por_isin: dict[str, str],
) -> tuple[DadosEvento | None, str]:  # fmt: skip
    if q < 0:
        return None, ""  # venda a descoberto na data com: sem proposta (o motor falha alto)
    # subscrição e bonificação podem entregar outra classe (quem tem PN recebe ON): o crédito é na
    # classe emitida, não na ação que o titular tem
    outra = (tipo in (TipoEvento.BONIFICACAO, TipoEvento.SUBSCRICAO)
             and ev.isin_emitido not in ("", ev.isin))  # fmt: skip
    emitida = por_isin.get(ev.isin_emitido) if outra else None
    classe = ""
    if outra:
        classe = (f"a companhia entrega {emitida}, outra classe da ação que você tem" if emitida
                  else "a companhia entrega outra classe, que o COIR não identificou: confira o "
                       "ativo")  # fmt: skip
    base = DadosEvento(
        tipo=tipo or TipoEvento.DESDOBRAMENTO, data=data_ex, conta=conta, ticker=emitida or ticker,
        cv="C", origem="b3",
        natureza=NaturezaEvento.OUTRA_CLASSE if emitida else None,
        observacao=f"pacote público da B3: {ev.tipo} {_fator(ev, tipo)} (data com "
                   f"{ev.data_com})",
    )  # fmt: skip
    if cisao:
        destino = por_isin.get(ev.isin_emitido)
        if destino is None:
            return None, ""
        recebida, _ = _fracao(q * fator)
        nota = f"informe o percentual do custo que vai para {destino} (aviso aos acionistas)"
        return replace(base, tipo=TipoEvento.CISAO, cv=None, quantidade=q, ticker_destino=destino,
                       quantidade_destino=recebida), nota  # fmt: skip
    if tipo is TipoEvento.GRUPAMENTO:
        quantidade, nota = _fracao(q * fator)
        if fator <= GRUPAMENTO_EXTREMO:
            nota = "; ".join(x for x in (AVISO_GRUPAMENTO_EXTREMO, nota) if x)
        return replace(base, cv="V", quantidade=q - quantidade), nota
    quantidade, nota = _fracao(q * fator)
    if tipo is TipoEvento.SUBSCRICAO:
        quantidade = quantidade.to_integral_value(ROUND_FLOOR)  # só se subscrevem ações inteiras
        preco = Decimal(ev.preco or 0)
        dia = date.fromisoformat(ev.data_subscricao) if ev.data_subscricao else data_ex
        return replace(base, data=dia, quantidade=quantidade,
                       valor=(quantidade * preco).quantize(Decimal("0.01"))), (
            "; ".join(x for x in (classe, f"subscrição é opcional (preço {preco}, período "
                                          f"{ev.periodo}): lance só o que você exerceu")
                      if x))  # fmt: skip
    if tipo is TipoEvento.BONIFICACAO:
        # o pacote traz o custo por ação quando o aviso aos acionistas informa: a proposta já sai
        # com o custo total; sem ele, o usuário informa (com custo zero o PM da venda fica errado)
        custo = Decimal(ev.custo_atribuido or 0)
        if custo:
            valor = (quantidade * custo).quantize(Decimal("0.01"), ROUND_HALF_UP)
            pedido = (f"custo atribuído de R$ {custo} por ação, do aviso aos acionistas "
                      "(confira)")  # fmt: skip
            nota = "; ".join(x for x in (classe, pedido, nota) if x)
            return replace(base, quantidade=quantidade, valor=valor), nota
        nota = "; ".join(x for x in (classe, "informe o custo atribuído (aviso aos acionistas)",
                                     nota) if x)  # fmt: skip
    return replace(base, quantidade=quantidade), nota


def _ajuste_serie(
    tipo: TipoEvento, serie: str, conta: int, q: Decimal, fator: Decimal, data_ex: date
) -> tuple[DadosEvento, str]:  # fmt: skip
    novo = abs(q) * (fator if tipo is TipoEvento.GRUPAMENTO else 1 + fator)
    delta = novo.to_integral_value(ROUND_FLOOR) - abs(q)
    cv = ("C" if delta > 0 else "V") if q > 0 else ("V" if delta > 0 else "C")
    return DadosEvento(tipo=tipo, data=data_ex, conta=conta, ticker=serie, cv=cv,
                       quantidade=abs(delta), origem="b3",
                       observacao="ajuste da série pelo evento do ativo-objeto"), (
        "série de opção aberta: a B3 ajusta a quantidade e o preço de exercício (e às vezes o "
        "código; se mudou, use Incorporação / ajuste de série)")  # fmt: skip
