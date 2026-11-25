# pyright: reportPrivateUsage=false, reportUnusedFunction=false
# (os nomes com _ são internos do pacote, usados entre os arquivos dele)
"""O percurso do livro dia a dia (eventos, termo, day trade, compras, exercícios, vendas,
vencimento) e o isolamento do ativo que o motor não consegue apurar."""

from calendar import monthrange
from collections import defaultdict
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import Decimal

from coir.domain.fiscal import (
    ZERO,
    AtivoIsolado,
    Categoria,
    Evento,
    Movimento,
    NotaMotor,
    OperacaoMotor,
    PosicaoAtivo,
    TipoAtivo,
    TipoEvento,
    Visao,
)
from coir.domain.parametros import ParametrosFiscais
from coir.engine.movimentos.day_trade import _day_trade
from coir.engine.movimentos.eventos import (
    _chave_evento,
    _cisao,
    _direito_nao_exercido,
    _dividendo_em_bens,
    _evento,
    _exercicio_direito,
    _incorporacao,
    _leilao_fracao,
)
from coir.engine.movimentos.livro import (
    MERCADOS_ACAO_A_VISTA,
    MERCADOS_DT,
    MERCADOS_EXERCICIO,
    SEM_IRRF_DA_VENDA,
    Chave,
    Livro,
    Meses,
    OpsDoDia,
    _avisar_come_cotas,
    _avisar_conjugadas,
    _categoria,
    _de,
    _Dia,
    _grupo,
    _grupo_opcao,
    _Op,
    _ordem,
    _ordem_negocio,
    _proporcao,
)
from coir.engine.movimentos.posicao import _aplicar, _exercicio, _fim_da_isencao, _lote_pme
from coir.engine.movimentos.termo import _avisar_termo_descoberto, _venda_a_termo
from coir.engine.posicao import Posicao
from coir.engine.rateio import ratear
from coir.erros import MotorError


def movimentar(
    notas: Iterable[NotaMotor],
    eventos: Iterable[Evento],
    parametros: ParametrosFiscais,
    ate: date,
    visao: Visao = Visao.TITULAR,
    avisos: list[str] | None = None,
    isolados: list[AtivoIsolado] | None = None,
) -> list[Movimento]:
    """Movimentos de posição até a data `ate` (inclusive) na visão pedida (a apuração usa só a
    por titular). Situações aceitas mas que pedem conferência (ex.: venda que passa de comprado a
    vendido) vão para `avisos`.

    Sem `isolados`, um ativo que o motor não consegue apurar para tudo (MotorError). Com a lista,
    esse ativo sai do cálculo a partir do dia do problema, o resto segue, e ele entra na lista
    (pendência para o usuário resolver, em geral um evento corporativo que falta)."""
    livro = _executar(notas, eventos, parametros, ate, visao, isolados)
    if avisos is not None:
        avisos += livro.avisos
    return livro.movimentos


def carteira(
    notas: Iterable[NotaMotor],
    eventos: Iterable[Evento],
    parametros: ParametrosFiscais,
    ate: date,
    visao: Visao = Visao.TITULAR,
    isolados: list[AtivoIsolado] | None = None,
) -> list[PosicaoAtivo]:
    """Posições abertas no fim do dia `ate`. Por titular é o custo médio fiscal (apuração e
    IRPF); por corretora e por conta é a visão operacional, que não entra no imposto. Compras a
    termo ainda não liquidadas não estão na carteira. Com `isolados` (ver `movimentar`), o ativo
    isolado fica fora da carteira: a posição dele não é conhecida."""
    livro = _executar(notas, eventos, parametros, ate, visao, isolados)
    saida: list[PosicaoAtivo] = []
    for (titular, grupo, ativo, vencimento), pos in livro.posicoes.items():
        if not pos.quantidade or (titular, ativo) in livro.cortes:
            continue
        conta = grupo if isinstance(grupo, int) else None
        corretora = grupo if isinstance(grupo, str) else None
        if conta is not None:
            corretora = livro.corretoras[conta]
        saida.append(
            PosicaoAtivo(
                titular=titular, corretora=corretora, conta=conta, ativo=ativo,
                ticker=livro.tickers[ativo], tipo=livro.tipos[ativo], vencimento=vencimento,
                quantidade=pos.quantidade, custo=pos.custo,
            )
        )  # fmt: skip
    return sorted(saida, key=lambda p: (p.titular, p.corretora or "", p.conta or 0, p.ticker,
                                        p.vencimento or date.min))  # fmt: skip


def _executar(
    notas: Iterable[NotaMotor],
    eventos: Iterable[Evento],
    parametros: ParametrosFiscais,
    ate: date,
    visao: Visao,
    isolados: list[AtivoIsolado] | None = None,
) -> Livro:
    """Percorre o livro; com `isolados`, refaz sem o ativo em que parou até fechar."""
    notas, eventos = list(notas), list(eventos)
    cortes: dict[tuple[int, int], tuple[date, str]] = {}
    while True:
        livro = Livro(visao)
        livro.cortes = {k: d for k, (d, _) in cortes.items()}
        try:
            meses = _percorrer(livro, notas, eventos, parametros, ate)
            _desde_zerado(livro, meses)
        except MotorError as erro:
            if isolados is None or livro.em_curso is None or erro.data is None:
                raise
            titular, ativos = livro.em_curso
            novos = {(titular, a) for a in ativos if (titular, a) not in cortes}
            if not novos:
                raise
            for chave in novos:
                cortes[chave] = (erro.data, str(erro))
            continue
        break
    if isolados is not None:
        for (titular, ativo), corte in sorted(livro.cortes.items(), key=lambda x: x[1]):
            motivo = _motivo(cortes, livro.puxados, titular, ativo)
            isolados.append(AtivoIsolado(
                titular=titular, ativo=ativo, ticker=livro.tickers.get(ativo, str(ativo)),
                desde=corte, motivo=motivo,
                categoria=_categoria(livro.tipos.get(ativo, TipoAtivo.ACAO), day_trade=False),
                meses=tuple(sorted(meses.get((titular, ativo), set()))),
            ))  # fmt: skip
    return livro


def _motivo(
    cortes: dict[tuple[int, int], tuple[date, str]], puxados: dict[tuple[int, int], int],
    titular: int, ativo: int,
) -> str:  # fmt: skip
    """Motivo do isolamento; o ativo puxado por outro ligado (exercício, incorporação) herda o
    motivo de quem o puxou, seguindo a cadeia até o ativo em que o motor parou."""
    visto: set[int] = set()
    while (titular, ativo) not in cortes and (titular, ativo) in puxados and ativo not in visto:
        visto.add(ativo)
        ativo = puxados[(titular, ativo)]
    _, motivo = cortes.get((titular, ativo), (None, ""))
    return motivo or next(m for (t, _), (_, m) in cortes.items() if t == titular)


def _desde_zerado(livro: Livro, meses: Meses) -> None:
    """O evento que falta pode ser anterior ao dia em que o motor parou: os meses do ativo desde
    a última vez que a posição zerou também ficam incompletos (critério do COIR de 03/11/2026,
    parecer do especialista)."""
    for (titular, ativo), corte in livro.cortes.items():
        movs = [m for m in livro.movimentos
                if m.titular == titular and m.ativo == ativo and m.data < corte]  # fmt: skip
        zerou = max((i for i, m in enumerate(movs) if m.quantidade_apos == 0), default=-1)
        meses[(titular, ativo)] |= {(m.data.year, m.data.month) for m in movs[zerou + 1 :]}


def _venda_isolada(livro: Livro, dia: date, o: _Op) -> None:
    """Venda de ativo isolado: o resultado fica de fora (o custo não é conhecido), mas o valor
    da venda é: entra na soma do limite de isenção do mês e na base que reparte o IRRF (decisão
    do usuário de 03/11/2026). O IRRF da nota continua crédito do mês. O ativo sem tipo conta
    como ação: se for, não contar isentaria vendas de ações acima do limite (imposto a menos, com
    multa de ofício); se não for, contar só faz pagar a mais, o que se restitui (parecer do
    especialista de 09/11/2026). A apuração avisa o mês em que isso decide
    o limite."""
    op = o.op
    if op.cv != "V" or op.mercado == "TERMO":
        return
    pme = op.pme_ate is not None and dia <= op.pme_ate
    acao = (op.tipo in (TipoAtivo.ACAO, TipoAtivo.UNIT, TipoAtivo.SEM_TIPO)
            and op.mercado in MERCADOS_ACAO_A_VISTA)  # fmt: skip
    base_irrf = op.total if (op.mercado not in MERCADOS_EXERCICIO and not pme
                             and op.tipo not in SEM_IRRF_DA_VENDA) else ZERO  # fmt: skip
    livro.movimentos.append(Movimento(
        data=dia, titular=o.nota.titular, conta=o.nota.conta, corretora=o.nota.corretora,
        ativo=op.ativo, ticker=op.ticker, origem=f"op:{op.id}",
        regra="ativo isolado (pendência): resultado fora do cálculo; o valor da venda conta no "
        "limite de isenção do mês",
        categoria=_categoria(op.tipo, day_trade=False),
        grupo=livro.grupos.get(op.ativo) or _grupo(op.tipo), tipo=op.tipo, cv="V",
        quantidade=op.quantidade, valor=o.valor, custos=o.custos,
        venda_acao_a_vista=op.total if acao and not pme else ZERO, base_irrf=base_irrf,
        quantidade_apos=ZERO, custo_apos=ZERO,
    ))  # fmt: skip


def _ativos_do_evento(e: Evento) -> tuple[int | None, ...]:
    """Os ativos que o evento liga no isolamento. O dividendo em bens só mexe no bem recebido: a
    posição na companhia que pagou não muda, e uma não deve puxar a outra."""
    if e.tipo is TipoEvento.DIVIDENDO_EM_BENS:
        return (e.ativo_destino,)
    return (e.ativo, e.ativo_destino)


def _isolar(
    livro: Livro, ops: OpsDoDia, eventos: list[Evento]
) -> tuple[OpsDoDia, OpsDoDia, list[Evento], Meses]:
    """Tira do livro as operações e os eventos de ativos isolados a partir do corte. O que está
    ligado vai junto: o exercício leva a série exercida; a incorporação e a cisão, o destino."""
    meses: Meses = defaultdict(set)
    if not livro.cortes:
        return ops, [], eventos, meses
    itens = [(dia, o.nota.titular, (o.op.ativo, o.op.opcao)) for dia, o in ops] + [
        (e.data, e.titular, _ativos_do_evento(e)) for e in eventos
    ]
    mudou = True
    while mudou:  # o corte de um ativo pode puxar o de outro
        mudou = False
        for dia, titular, ativos in itens:
            ligados = [a for a in ativos if a is not None]
            fora = [a for a in ligados if livro.fora(titular, a, dia)]
            if fora:
                for a in ligados:
                    if not livro.fora(titular, a, dia):
                        livro.cortes[(titular, a)] = dia
                        livro.puxados[(titular, a)] = fora[0]
                        mudou = True

    def tirar(dia: date, titular: int, ativos: tuple[int | None, ...]) -> bool:
        ligados = [a for a in ativos if a is not None]
        if not any(livro.fora(titular, a, dia) for a in ligados):
            return False
        for a in ligados:
            meses[(titular, a)].add((dia.year, dia.month))
        return True

    fica: OpsDoDia = []
    tiradas: OpsDoDia = []
    for dia, o in ops:
        (tiradas if tirar(dia, o.nota.titular, (o.op.ativo, o.op.opcao)) else fica).append((dia, o))
    eventos = [e for e in eventos if not tirar(e.data, e.titular, _ativos_do_evento(e))]
    return fica, tiradas, eventos, meses


@dataclass
class _Lidas:
    """O que a leitura das notas deixa para o percurso."""

    ops: list[tuple[date, _Op]] = field(default_factory=list[tuple[date, _Op]])  # dia de entrada
    vencimentos: set[date] = field(default_factory=set[date])
    vendas_a_termo: list[tuple[NotaMotor, OperacaoMotor]] = field(
        default_factory=list[tuple[NotaMotor, OperacaoMotor]]
    )
    # compra a termo com aviso
    compras_sem_b3: list[tuple[NotaMotor, OperacaoMotor]] = field(
        default_factory=list[tuple[NotaMotor, OperacaoMotor]]
    )
    # titular, corretora, ativo, pregão
    compras_a_vista: set[tuple[int, str, int, date]] = field(
        default_factory=set[tuple[int, str, int, date]]
    )


def _ler_notas(
    livro: Livro, notas: list[NotaMotor], parametros: ParametrosFiscais, ate: date
) -> _Lidas:
    """As operações das notas até `ate`, com o dia em que cada uma entra no livro (a termo: na
    liquidação), os vencimentos de opção e o que os avisos de termo usam; registra no livro o
    ticker, o tipo e o grupo de cada ativo."""
    lidas = _Lidas()
    for nota in notas:
        if nota.data > ate or not nota.operacoes:
            continue  # nota BM&F: só ajustes (futuros.py)
        if nota.ajustes:
            raise MotorError(f"nota {nota.numero}: operações Bovespa e ajustes BM&F juntos")
        custos = ratear(nota.custos, [o.total for o in nota.operacoes])
        for op, custo in zip(nota.operacoes, custos, strict=True):
            livro.tickers[op.ativo] = op.ticker
            livro.tipos[op.ativo] = op.tipo
            if op.tipo is TipoAtivo.OPCAO:
                livro.grupos[op.ativo] = _grupo_opcao(op, parametros, nota.data)
                aviso = (f"Opção {op.ticker}: o COIR não identificou o ativo-objeto pelo código. "
                         "No IRPF ela entra na linha \"Mercado de opções - ações\", o certo para "
                         "opção de ação; só se for opção de ETF ou índice é outra linha (avise "
                         "pelo formulário do site para corrigirmos a tabela).")  # fmt: skip
                if (op.objeto_tipo is None and livro.visao is Visao.TITULAR
                        and aviso not in livro.avisos):  # fmt: skip
                    livro.avisos.append(aviso)
            if op.tipo is TipoAtivo.FIP and op.cv == "V" and livro.visao is Visao.TITULAR:
                _avisar_come_cotas(livro, op.ticker, nota.data, parametros)
            item = _Op(nota, op, custo)
            if op.mercado == "TERMO":
                if op.liquidacao_termo is None:
                    raise MotorError(f"nota {nota.numero} de {nota.data}, {op.ticker}: termo sem "
                                     "data de liquidação (prazo)")  # fmt: skip
                if op.cv == "V":
                    lidas.vendas_a_termo.append((nota, op))
                elif op.aviso:
                    lidas.compras_sem_b3.append((nota, op))
                if op.liquidacao_termo <= ate:
                    lidas.ops.append((op.liquidacao_termo, item))
                    if (op.aviso and livro.visao is Visao.TITULAR
                            and op.aviso not in livro.avisos):  # fmt: skip
                        livro.avisos.append(op.aviso)
                continue
            lidas.ops.append((nota.data, item))
            if op.cv == "C" and op.mercado in MERCADOS_ACAO_A_VISTA:
                lidas.compras_a_vista.add((nota.titular, nota.corretora, op.ativo, nota.data))
            if op.tipo is TipoAtivo.OPCAO:
                if op.vencimento is None or op.vencimento < nota.data:
                    raise MotorError(
                        f"opção {op.ticker} em {nota.data}: vencimento {op.vencimento} ausente ou "
                        "anterior ao negócio (confira data/vencimentos_opcoes*.csv)"
                    )
                lidas.vencimentos.add(op.vencimento)
    return lidas


def _eventos_ate(livro: Livro, eventos: list[Evento], ate: date) -> list[Evento]:
    """Os eventos até `ate`, registrando no livro o ticker e o tipo do ativo (e do destino)."""
    ate_hoje: list[Evento] = []
    for evento in eventos:
        if evento.data <= ate:
            livro.tickers.setdefault(evento.ativo, evento.ticker)
            livro.tipos.setdefault(evento.ativo, evento.tipo_ativo)
            if evento.ativo_destino is not None and evento.ticker_destino is not None:
                livro.tickers.setdefault(evento.ativo_destino, evento.ticker_destino)
                livro.tipos.setdefault(evento.ativo_destino,
                                       evento.tipo_ativo_destino or evento.tipo_ativo)  # fmt: skip
            ate_hoje.append(evento)
    return ate_hoje


def _percorrer(
    livro: Livro, notas: list[NotaMotor], eventos: list[Evento], parametros: ParametrosFiscais,
    ate: date,
) -> Meses:  # fmt: skip
    """Monta os dias e processa; devolve os meses com itens tirados por ativo isolado."""
    dias: dict[date, _Dia] = defaultdict(_Dia)
    lidas = _ler_notas(livro, notas, parametros, ate)
    if livro.visao is Visao.TITULAR:
        _avisar_conjugadas(livro, lidas.vendas_a_termo, lidas.compras_a_vista)
    ate_hoje = _eventos_ate(livro, eventos, ate)
    ops, tiradas, ate_hoje, meses = _isolar(livro, lidas.ops, ate_hoje)
    livro.parametros = parametros
    fins = {x.pme_ate + timedelta(days=1)
            for x in (*(o.op for _, o in ops), *ate_hoje) if x.pme_ate is not None}  # fmt: skip
    livro.fins_pme = {d for d in fins if d <= ate}
    for dia in livro.fins_pme:
        dias.setdefault(dia, _Dia())
    for dia, item in ops:
        (dias[dia].termos if item.op.mercado == "TERMO" else dias[dia].operacoes).append(item)
    for evento in ate_hoje:
        dias[evento.data].eventos.append(evento)
    for venc in lidas.vencimentos:
        if venc <= ate:
            dias.setdefault(venc, _Dia())  # o pó acontece mesmo sem negócio no dia
    ofertas = [e for e in eventos if e.tipo is TipoEvento.OFERTA_PUBLICA
               and not livro.fora(e.titular, e.ativo, e.data)]  # fmt: skip
    for dia in sorted(dias):
        livro.hoje = dia
        try:
            _antecipar_ofertas(livro, dia, dias, ofertas, parametros)
            _processar_dia(livro, dia, dias[dia], parametros)
        except MotorError as erro:
            erro.data = erro.data or dia
            raise
    livro.em_curso = None
    _avisar_suspeitas(livro)
    if livro.visao is Visao.TITULAR:
        _avisar_termo_descoberto(livro, lidas.compras_sem_b3)
    for dia, o in tiradas:
        _venda_isolada(livro, dia, o)
    return meses


def _antecipar_ofertas(
    livro: Livro, dia: date, dias: dict[date, _Dia], ofertas: list[Evento],
    parametros: ParametrosFiscais,
) -> None:  # fmt: skip
    """A negociação pode começar antes da liquidação da oferta: a venda do dia que deixaria a
    posição vendida é das ações da oferta, então a aquisição entra antes dela, e o ganho fica no
    mês do pregão da venda (critério do COIR de 08/11/2026; IN RFB 1.585/2015 art. 58). Antecipa
    só o que cobre a venda, da oferta mais próxima para a mais distante. A posição vendida há mais
    tempo que a janela não é venda das ações da oferta: o motor falha alto em _evento."""
    hoje = dias[dia]
    if not ofertas or not hoje.operacoes:
        return
    janela = timedelta(days=int(parametros.decimal("janela_liquidacao_oferta_dias", dia)))
    for e in sorted(ofertas, key=lambda x: (x.data, x.id)):
        if not dia < e.data <= dia + janela:
            continue
        chave = _chave_evento(livro, e)
        saldo = livro.posicoes[chave].quantidade
        for x in hoje.eventos:  # os eventos do dia vêm antes das operações
            if x.tipo is TipoEvento.OFERTA_PUBLICA and _chave_evento(livro, x) == chave:
                saldo += x.quantidade or ZERO
        for o in (*hoje.termos, *hoje.operacoes):
            if livro.chave(o.nota.titular, _de(o), o.op.ativo) == chave:
                saldo += o.op.quantidade if o.op.cv == "C" else -o.op.quantidade
        if saldo >= 0:
            continue
        ofertas.remove(e)
        if e.data in dias:
            dias[e.data].eventos.remove(e)
        hoje.eventos.append(e.model_copy(update={"data": dia}))
        if livro.visao is Visao.TITULAR:
            livro.avisos.append(
                f"Oferta de {e.ticker}: venda em {dia:%d/%m/%Y}, antes da liquidação em "
                f"{e.data:%d/%m/%Y}. A compra entra no dia da venda, para o ganho ficar nesse mês."
            )


def _avisar_suspeitas(livro: Livro) -> None:
    """Avisa a venda que passou de comprado a vendido quando a posição continua vendida no fim
    do mês (recompra logo depois é venda a descoberto comum, sem aviso)."""
    for chave, dia, quantidade, antes in livro.suspeitas:
        titular, _, ativo, _ = chave
        fim = date(dia.year, dia.month, monthrange(dia.year, dia.month)[1])
        apos = [m.quantidade_apos for m in livro.movimentos
                if (m.titular, m.ativo) == (titular, ativo) and m.data <= fim]  # fmt: skip
        if apos and apos[-1] < 0:
            aviso = (
                f"{dia}: venda de {quantidade} {livro.tickers[ativo]} maior que a posição "
                f"({antes}), que ficou vendida no fim do mês ({apos[-1]}). Pode faltar uma "
                "nota de compra (de outra corretora ou anterior às notas importadas), a custódia "
                "inicial (Cadastros) ou um evento corporativo (tela Eventos)"
            )
            livro.avisos.append(aviso)


def _direitos_recebidos(livro: Livro, dia: date, ops: Sequence[_Op]) -> None:
    """Direito de subscrição vendido sem nota de compra: foi recebido da companhia ou do fundo,
    sem pagar nada. A parte da venda que passa da posição (e das compras do dia) entra antes,
    com custo zero (IN RFB 1.585/2015 art. 58 §7º, por analogia: a norma não trata do direito
    recebido; parecer do especialista de 09/11/2026), em vez de virar venda a descoberto sem
    resultado. Posição já vendida fica como está: aí a regra não se aplica."""
    liquido: dict[Chave, Decimal] = defaultdict(lambda: ZERO)
    venda: dict[Chave, _Op] = {}
    for o in ops:
        if o.op.tipo is not TipoAtivo.DIREITO or o.op.mercado not in MERCADOS_DT:
            continue
        chave = livro.chave(o.nota.titular, _de(o), o.op.ativo)
        liquido[chave] += o.op.quantidade if o.op.cv == "C" else -o.op.quantidade
        if o.op.cv == "V":
            venda.setdefault(chave, o)
    for chave, saldo in liquido.items():
        pos = livro.posicoes[chave]
        falta = -(pos.quantidade + saldo)
        if pos.quantidade < 0 or falta <= 0:
            continue
        o = venda[chave]
        livro.processando(o.nota.titular, o.op.ativo)
        if pos.custo > 0 and livro.visao is Visao.TITULAR:
            # com direitos comprados na posição, a data do crédito muda o custo médio: lançado só
            # na venda, o ganho desta venda sai menor (ou adiado) do que o devido
            livro.avisos.append(
                f"{dia:%d/%m/%Y}: {o.op.ticker}: direito vendido sem o crédito lançado: os "
                f"{falta} recebidos entraram com custo zero só nesta venda. Como havia direitos "
                "comprados, o custo médio pode ter ficado maior e o ganho menor: lance o "
                '"Direito recebido" na data do crédito (tela Eventos; a movimentação da B3 '
                "traz a data)"
            )
        pos.comprar(falta, ZERO)
        livro.registrar(chave, dia, f"direito:{o.op.id}",
                        "direito recebido sem custo: entra antes da venda, com custo zero (IN RFB "
                        "1.585/2015 art. 58 §7º, por analogia)", de=_de(o),
                        categoria=Categoria.COMUM, cv="C", quantidade=falta,
                        valor=ZERO)  # fmt: skip


def _processar_dia(livro: Livro, dia: date, d: _Dia, params: ParametrosFiscais) -> None:
    if dia in livro.fins_pme:
        _fim_da_isencao(livro, dia)
    # cisões depois dos demais eventos do dia (agrupadas por posição de origem e destino)
    cisoes: dict[tuple[Chave, int | None], list[Evento]] = defaultdict(list)
    for evento in sorted(d.eventos, key=lambda e: e.id):
        if evento.tipo is TipoEvento.DIVIDENDO_EM_BENS:
            # só o bem recebido muda; a companhia que pagou não entra na conta
            livro.processando(evento.titular, evento.ativo_destino)
            _dividendo_em_bens(livro, evento, params)
            continue
        livro.processando(evento.titular, evento.ativo, evento.ativo_destino)
        if evento.tipo is TipoEvento.CISAO:
            # as contas que formam a posição (visão por titular) entram juntas: o custo repartido
            # é o de antes do evento, não o que sobra depois da primeira conta
            cisoes[(_chave_evento(livro, evento), evento.ativo_destino)].append(evento)
        elif evento.tipo is TipoEvento.INCORPORACAO:
            _incorporacao(livro, evento, params)
        elif evento.tipo is TipoEvento.LEILAO_FRACAO:
            _leilao_fracao(livro, evento, params)
        elif evento.tipo is TipoEvento.DIREITO_NAO_EXERCIDO:
            _direito_nao_exercido(livro, evento)
        elif evento.tipo is TipoEvento.EXERCICIO_DIREITO:
            _exercicio_direito(livro, evento)
        else:
            _evento(livro, evento)
    origens = [chave for chave, _ in cisoes]
    for (chave, _), grupo in cisoes.items():
        livro.processando(chave[0], chave[2], grupo[0].ativo_destino)
        if origens.count(chave) > 1:  # o percentual é sobre o custo de antes de todas
            raise MotorError(f"cisão de {grupo[0].ticker} ({grupo[0].data}) para dois ativos no "
                             "mesmo dia (não implementado)")  # fmt: skip
        _cisao(livro, chave, grupo)
    livro.termo_hoje = {}
    # as compras a termo liquidam antes das vendas a termo: as ações recebidas podem ser entregues
    for t in sorted(d.termos, key=lambda t: (t.op.cv == "V", t.nota.id, t.op.ordem)):
        livro.processando(t.nota.titular, t.op.ativo)
        chave = livro.chave(t.nota.titular, _de(t), t.op.ativo)
        pos = livro.posicoes[chave]
        if t.op.cv == "V":
            _venda_a_termo(livro, dia, chave, t)
            continue
        # ações vendidas antes da liquidação: a posição ficou vendida e o termo a cobre (resultado
        # na liquidação, limite de isenção do mês da venda, como na recompra da venda a descoberto)
        mes_limite = pos.descoberto_desde if pos.quantidade < 0 else None
        antes = pos.quantidade
        baixa = pos.comprar(t.op.quantidade, t.valor)
        _lote_pme(pos, t.op, dia, t.op.quantidade, t.valor, antes)
        # só a parte que ficou comprada pode ser vendida nesta data pelo preço a termo
        longo = t.op.quantidade if antes >= 0 else max(ZERO, pos.quantidade)
        if longo:
            lote = livro.termo_hoje.setdefault(chave, Posicao())
            lote.comprar(longo, _proporcao(t.valor, longo, t.op.quantidade))
        if pos.quantidade >= 0:
            pos.descoberto_desde = None
        regra = "termo liquidado: entra pelo preço a termo + custos da compra"
        if baixa.resultado or baixa.custo_baixa:
            regra += "; cobre a venda feita antes da liquidação (resultado nesta data)"
        # o ganho na liquidação do termo nunca é isento (IN RFB 1.585/2015 art. 59 §2º IV)
        livro.registrar(chave, dia, f"termo:{t.op.id}", regra, de=_de(t),
                        categoria=Categoria.COMUM, cv="C", quantidade=t.op.quantidade,
                        valor=t.valor, custos=t.custos, baixa=baixa, isentavel=False,
                        mes_limite=mes_limite)  # fmt: skip
    if params.booleano("direito_recebido_custo_zero", dia):
        _direitos_recebidos(livro, dia, d.operacoes)
    mesma_corretora = params.booleano("day_trade_mesma_corretora", dia)
    pela_marcacao = params.texto("day_trade_criterio", dia) == "marcacao_corretora"
    livro.em_curso = None
    compras, vendas, exercicios = _day_trade(livro, dia, d.operacoes, mesma_corretora,
                                             pela_marcacao)  # fmt: skip
    for p in sorted(compras, key=_ordem):
        livro.processando(p.base.nota.titular, p.base.op.ativo)
        _aplicar(livro, dia, p)
    for o in sorted(exercicios, key=_ordem_negocio):
        livro.processando(o.nota.titular, o.op.ativo, o.op.opcao)
        _exercicio(livro, dia, o)
    for p in sorted(vendas, key=_ordem):
        livro.processando(p.base.nota.titular, p.base.op.ativo)
        _aplicar(livro, dia, p)
    for chave, pos in list(livro.posicoes.items()):
        if pos.quantidade and chave[3] == dia and not livro.fora(chave[0], chave[2], dia):
            livro.processando(chave[0], chave[2])
            quantidade = pos.quantidade
            de = livro.origem_unica(chave)
            baixa = pos.virar_po()
            livro.registrar(chave, dia, "vencimento",
                            "opção não exercida no vencimento (pó): "
                            + ("titular perde o prêmio" if quantidade > 0
                               else "lançador ganha o prêmio"), de=de,
                            categoria=Categoria.COMUM, cv=None, quantidade=quantidade,
                            valor=ZERO, baixa=baixa)  # fmt: skip
