"""Eventos sugeridos para a tela Eventos: os da movimentação da B3 (recibos, bonificações,
amortizações) juntos com os do pacote público, e os avisos ligados a eles."""

import re
from collections import defaultdict
from collections.abc import Sequence
from dataclasses import replace
from datetime import date, datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from sqlalchemy import func, select
from sqlalchemy.orm import Session, selectinload

from coir import auditoria, fontes
from coir.db import repositorio as repo
from coir.db.modelos import (
    Conta,
    EventoCorporativo,
    MovimentacaoB3,
    NotaCorretagem,
    PropostaIgnorada,
)
from coir.db.sessao import conectar
from coir.domain.fiscal import (
    NaturezaEvento,
    TipoAtivo,
    TipoEvento,
)
from coir.domain.parametros import ReferenciasFiscais
from coir.domain.referencias import RE_SERIE_OPCAO, Referencias
from coir.engine.calendario import proximo_pregao
from coir.erros import ImportacaoError, ReferenciaError
from coir.servicos import trilha
from coir.servicos.apuracao import Entradas
from coir.servicos.eventos.dados import (
    JANELA,
    NOTA_DATA_EX,
    DadosEvento,
    EventoLancado,
    Proposta,
    _ativos_do_ticker,
    _ja_lancado,
    listar,
)
from coir.servicos.eventos.pacote import _do_pacote

# movimentação da B3 (minúsculas) -> tipo e C/V do evento proposto
PROPOSTAS: dict[str, tuple[TipoEvento, str | None]] = {
    "desdobro": (TipoEvento.DESDOBRAMENTO, "C"),
    "bonificação em ativos": (TipoEvento.BONIFICACAO, "C"),
    "restituição de capital": (TipoEvento.REDUCAO_CAPITAL, "V"),
    # com código de recibo (final 9 ou 10), vira o exercício do direito (ver _mapa_proposta)
    "recibo de subscrição": (TipoEvento.SUBSCRICAO, "C"),
    "direito de subscrição": (TipoEvento.DIREITO_RECEBIDO, "C"),
    "incorporação": (TipoEvento.INCORPORACAO, None),
    "cisão": (TipoEvento.CISAO, None),
    "leilão de fração": (TipoEvento.LEILAO_FRACAO, None),
    # só em cotas de FI-Infra, FIP-IE e FIP comum (ver _mapa_proposta)
    "amortização": (TipoEvento.REDUCAO_CAPITAL, "V"),
}


SEM_CUSTO_NO_PACOTE = (
    "o pacote público não traz o custo atribuído desta bonificação: informe-o (aviso aos "
    "acionistas); sem ele, o preço médio da venda fica errado"
)


AVISO_SEM_CUSTO = "informe o custo atribuído (aviso aos acionistas)"


NOTAS_PROPOSTA = {
    TipoEvento.BONIFICACAO: AVISO_SEM_CUSTO,
    TipoEvento.SUBSCRICAO: "ações subscritas: informe o valor pago",
    TipoEvento.INCORPORACAO: "a B3 mostra o ativo recebido: informe o de origem e a quantidade",
    TipoEvento.CISAO: "informe o ativo cindido e o percentual do custo",
    TipoEvento.LEILAO_FRACAO: "fração vendida pela companhia: venda comum no mês do crédito",
    TipoEvento.DIREITO_NAO_EXERCIDO: "o custo pago pelos direitos comprados vira perda em "
                                     "operações comuns; os recebidos sem custo não dão resultado",
    TipoEvento.DIREITO_RECEBIDO: "direito creditado pela companhia ou pelo fundo: entra com custo "
                                 "zero e compõe o custo médio dos direitos",
    TipoEvento.EXERCICIO_DIREITO: "direitos exercidos: informe o valor pago na subscrição; o custo "
                                  "dos direitos comprados soma ao do recibo",
}  # fmt: skip


NOTA_CREDITO = "data do crédito na B3"


NOTA_CONVERSAO = ("o recibo de subscrição virou ação: o custo passa inteiro, sem resultado; a data "
                  "é a do crédito das ações na B3")  # fmt: skip


JANELA_RECIBO = 180  # dias entre o crédito do recibo e o das ações (homologação do aumento)


ACAO_DO_RECIBO = {"9": "3", "10": "4"}  # recibo ON -> ação ON; recibo PN -> ação PN


DIREITO_DO_RECIBO = {"9": "1", "10": "2"}


NAO_EXERCIDO = ("não exercid", "nao exercid")  # "Direitos de Subscrição - Não Exercido"


AMORTIZACAO_FUNDO = ("amortização de cotas de FI-Infra ou FIP-IE: reduz o custo das cotas (Lei "
                     "11.478/2007 art. 2º §2º e IN RFB 1.585/2015 art. 33, FIP-IE; Lei 12.431/2011 "
                     "art. 3º e IN art. 34, FI-Infra); o que passar do "
                     "custo é rendimento (confira o informe do administrador)")  # fmt: skip


AMORTIZACAO_FIP = ("amortização de cotas de FIP: troque o valor pela parcela de custo que o "
                   "informe do administrador mostra. Desde 2024 só essa parcela devolve custo; o "
                   "restante é rendimento com IRRF de 15% (Lei 14.754/2023 art. 24). Até 2023 a "
                   "amortização reduz o custo e só o excedente é rendimento (IN RFB 1.585/2015 "
                   "art. 32 §2º)")  # fmt: skip


AMORTIZACAO_FII = ("amortização de cotas de FII: reduz o custo das cotas. Não há norma expressa; "
                   "a analogia é com a IN RFB 1.585/2015 arts. 16 §2º e 37 §3º, critério "
                   "conservador do COIR: o custo fica menor e o imposto na venda, maior. Se o "
                   "informe do administrador tratar parte do valor como rendimento, lance só a "
                   "parte de capital")  # fmt: skip


AMORTIZACAO_ETF = ("amortização de cotas de ETF: troque o valor pela parcela de custo que o "
                   "informe do administrador mostra. Desde 2024 só essa parcela devolve custo; o "
                   "restante é rendimento com IRRF de 15% (Lei 14.754/2023 arts. 18 II e "
                   "24)")  # fmt: skip


NOTA_AMORTIZACAO = {TipoAtivo.FIP: AMORTIZACAO_FIP, TipoAtivo.FII: AMORTIZACAO_FII,
                    TipoAtivo.ETF: AMORTIZACAO_ETF}  # fmt: skip


RE_RECIBO = re.compile(r"^(?P<raiz>[A-Z0-9]{4})(?P<classe>9|10)$")  # recibo de subscrição ON/PN


@auditoria.operacao("propostas")
def propostas(
    banco: Path, titular: int, refs: Referencias, fiscais: ReferenciasFiscais | None = None,
    pasta: Path | None = None, base: Entradas | None = None,
) -> list[Proposta]:  # fmt: skip
    """Eventos sugeridos, mais recentes primeiro: os do pacote público de referências que atingem
    a posição (com ``fiscais`` e ``pasta``; já com a data ex) e os da movimentação da B3 (data do
    crédito: a tela pede a data ex antes de gravar). O mesmo evento nas duas fontes aparece uma vez
    (a do pacote)."""
    with conectar(banco) as fabrica, fabrica() as s:
        ignoradas = set(s.scalars(select(PropostaIgnorada.chave)
                                  .where(PropostaIgnorada.titular_id == titular)))  # fmt: skip
        contas = list(s.scalars(select(Conta).where(Conta.titular_id == titular)
                                .options(selectinload(Conta.corretora))))  # fmt: skip
        # a compra numa oferta não é o evento proposto: não esconde uma bonificação de dias depois
        eventos_lancados = list(s.scalars(
            select(EventoCorporativo).where(EventoCorporativo.conta_id.in_([c.id for c in contas]),
                                            EventoCorporativo.tipo != TipoEvento.OFERTA_PUBLICA)
        ))  # fmt: skip
        lancados = [(e.ativo_id, e.ativo_destino_id, e.data) for e in eventos_lancados]
        saida: list[Proposta] = []
        bonificacoes = _bonificacoes_do_pacote(fiscais, pasta)
        linhas = list(s.scalars(
            select(MovimentacaoB3).where(MovimentacaoB3.titular_id == titular)
            .order_by(MovimentacaoB3.data.desc(), MovimentacaoB3.id)
        ))  # fmt: skip
        convertidos = _recibos_convertidos(linhas)
        for m in linhas:
            codigo = _codigo_b3(m)
            corretora = refs.corretora_b3(m.instituicao)
            das_contas = [c for c in contas if c.corretora.nome == corretora]
            conta = das_contas[0].id if len(das_contas) == 1 else None
            convertido = convertidos.get(m.chave)
            mapa = ((TipoEvento.INCORPORACAO, None) if convertido
                    else _mapa_proposta(m, codigo, fiscais))  # fmt: skip
            if mapa is None:
                continue
            tipo, cv = mapa
            if convertido or tipo is TipoEvento.EXERCICIO_DIREITO:
                saida.append(_proposta_do_recibo(s, refs, m, codigo, conta, corretora, convertido,
                                                 ignoradas, eventos_lancados))  # fmt: skip
                continue
            ticker = codigo
            recebido = tipo is TipoEvento.INCORPORACAO  # a B3 mostra o crédito de B
            # restituição de capital e amortização: a quantidade da B3 é a posição sobre a qual o
            # dinheiro foi pago, não ações ou cotas canceladas; levada ao evento, zeraria a posição
            so_valor = tipo is TipoEvento.REDUCAO_CAPITAL
            data, nota_data = m.data, NOTA_DATA_EX
            if tipo is TipoEvento.DIREITO_NAO_EXERCIDO:
                data, nota_data = _fim_do_prazo(codigo, m.data, fiscais, pasta)
            elif tipo is TipoEvento.DIREITO_RECEBIDO:
                nota_data = NOTA_CREDITO
            elif m.movimentacao.strip().lower() == "amortização":
                tipo_fundo = fiscais.tipos.tipo_ou_sem(codigo) if fiscais else None
                nota_data = NOTA_AMORTIZACAO.get(tipo_fundo, AMORTIZACAO_FUNDO)  # type: ignore[arg-type]
            dados = DadosEvento(
                tipo=tipo, data=data, conta=conta, ticker="" if recebido else ticker, cv=cv,
                quantidade=None if recebido or so_valor else m.quantidade,
                valor=(m.valor or Decimal(0)) if tipo in (TipoEvento.REDUCAO_CAPITAL,
                                                         TipoEvento.LEILAO_FRACAO) else Decimal(0),
                ticker_destino=ticker if recebido else None,
                quantidade_destino=m.quantidade if recebido else None,
                natureza=NaturezaEvento.INCORPORACAO_SOCIEDADE if recebido else None,
                observacao=f"proposto da movimentação da B3 ({m.movimentacao})", origem="b3",
            )  # fmt: skip
            if m.chave in ignoradas:
                situacao = "ignorada"
            elif _ja_lancado(s, refs, ticker, data, lancados):
                situacao = "lançado"
            else:
                situacao = "pendente"
            pedido = NOTAS_PROPOSTA.get(tipo, "")
            custo, referencia = None, ""
            if tipo is TipoEvento.BONIFICACAO and not dados.valor and dados.quantidade:
                achado = _custo_da_bonificacao(bonificacoes, ticker, m.data)
                if achado is not None:
                    custo, referencia = achado
                    dados = replace(dados, valor=(dados.quantidade * custo).quantize(
                        Decimal("0.01"), ROUND_HALF_UP))  # fmt: skip
                    pedido = (f"custo atribuído de R$ {custo} por ação, do pacote público "
                              "(aviso aos acionistas; confira)")  # fmt: skip
                else:
                    pedido = SEM_CUSTO_NO_PACOTE
            nota = "; ".join(x for x in (pedido, nota_data) if x)
            saida.append(Proposta(m.chave, m.data, m.movimentacao, codigo, m.quantidade,
                                  m.valor, corretora, situacao, dados, nota,
                                  custo_por_acao=custo, referencia=referencia))  # fmt: skip
        if fiscais is None or pasta is None:
            trilha.propostas(titular, saida)
            return _sem_pendencia_antes(saida)
        do_pacote = _do_pacote(s, titular, refs, fiscais, pasta, contas, ignoradas, lancados,
                               base)  # fmt: skip
        ultima = s.scalar(select(func.max(NotaCorretagem.data_pregao)).join(Conta)
                          .where(Conta.titular_id == titular))  # fmt: skip
    if ultima is not None:
        do_pacote = [replace(p, depois_da_ultima_nota=True, segura=False) if p.data > ultima else p
                     for p in do_pacote]  # fmt: skip
    perto = timedelta(days=JANELA)

    def mesma(p: Proposta, m: Proposta) -> bool:
        janela = (timedelta(days=JANELA_RESGATE) if p.dados.tipo is TipoEvento.REDUCAO_CAPITAL
                  else perto)  # fmt: skip
        return (p.dados.tipo is m.dados.tipo and p.codigo == m.codigo
                and abs(p.data - m.data) <= janela)  # fmt: skip

    # ignorou a proposta da movimentação: a do pacote para o mesmo evento também fica ignorada
    do_pacote = [replace(p, situacao="ignorada") if p.situacao == "pendente" and any(
        mesma(p, m) and m.situacao == "ignorada" for m in saida) else p
        for p in do_pacote]  # fmt: skip
    repetidas = [m for m in saida if any(mesma(p, m) for p in do_pacote)]
    lista = sorted([*do_pacote, *(m for m in saida if m not in repetidas)],
                   key=lambda p: p.data, reverse=True)  # fmt: skip
    trilha.propostas(titular, lista)
    return _sem_pendencia_antes(lista)


def _codigo_b3(m: MovimentacaoB3) -> str:
    return m.codigo or m.produto.split(" - ")[0].strip()


def _recibos_convertidos(linhas: Sequence[MovimentacaoB3]) -> dict[str, str]:
    """Chave da linha -> código do recibo de subscrição que virou as ações dela: uma "Atualização"
    de entrada na ação, de um recibo da mesma companhia e classe creditado até JANELA_RECIBO dias
    antes. A quantidade pode ser menor que a do recibo (parte dele pode ter sido vendida), e cada
    recibo só cobre conversões até a quantidade dele: depois disso, a "Atualização" é outra coisa.
    Vale o recibo mais recente que ainda cobre a linha."""
    recibos = [m for m in linhas if m.entrada and RE_RECIBO.match(_codigo_b3(m))
               and m.movimentacao.strip().lower() == "recibo de subscrição"]  # fmt: skip
    resta = {r.chave: r.quantidade for r in recibos}
    saida: dict[str, str] = {}
    for m in sorted(linhas, key=lambda x: (x.data, x.id)):
        if not m.entrada or m.movimentacao.strip().lower() != "atualização" or m.quantidade <= 0:
            continue
        codigo = _codigo_b3(m)
        candidatos = [
            r for r in recibos
            if (recibo := RE_RECIBO.match(_codigo_b3(r)))
            and codigo == recibo["raiz"] + ACAO_DO_RECIBO[recibo["classe"]]
            and m.quantidade <= resta[r.chave]
            and 0 <= (m.data - r.data).days <= JANELA_RECIBO
        ]  # fmt: skip
        if candidatos:
            usado = max(candidatos, key=lambda r: r.data)
            resta[usado.chave] -= m.quantidade
            saida[m.chave] = _codigo_b3(usado)
    return saida


def _lancado_do_recibo(
    s: Session, dia: date, lancados: Sequence[EventoCorporativo], *,
    exercicio: tuple[str, str, str] | None = None, conversao_de: str | None = None,
) -> bool:  # fmt: skip
    """O evento certo já está lançado perto da data? Olha o tipo, não só o ativo: o crédito do
    direito dias antes não é o exercício, e o exercício não é a conversão. Exercício: evento de
    exercício no direito ou no recibo, ou a subscrição do jeito antigo, direto na ação. Conversão:
    incorporação de natureza recibo saindo do recibo."""
    perto = timedelta(days=JANELA)

    def ativos(ticker: str) -> set[int]:
        achado = repo.ativo_existente(s, ticker, dia)
        return {achado} if achado is not None else _ativos_do_ticker(s, ticker)

    proximos = [e for e in lancados if abs(e.data - dia) <= perto]
    if conversao_de is not None:
        recibo = ativos(conversao_de)
        return any(
            e.tipo == TipoEvento.INCORPORACAO and e.ativo_id in recibo
            and e.natureza == NaturezaEvento.RECIBO_SUBSCRICAO
            for e in proximos
        )  # fmt: skip
    if exercicio is None:
        return False
    direito, recibo, acao = (ativos(t) for t in exercicio)
    return any(
        (e.tipo == TipoEvento.EXERCICIO_DIREITO
         and (e.ativo_id in direito or e.ativo_destino_id in recibo))
        or (e.tipo == TipoEvento.SUBSCRICAO and e.ativo_id in acao)
        for e in proximos
    )  # fmt: skip


def _proposta_do_recibo(
    s: Session, refs: Referencias, m: MovimentacaoB3, codigo: str, conta: int | None,
    corretora: str | None, convertido: str | None, ignoradas: set[str],
    lancados: Sequence[EventoCorporativo],
) -> Proposta:  # fmt: skip
    """O crédito do recibo propõe o exercício do direito (o direito sai e o recibo entra pelo
    valor pago mais o custo dos direitos); o crédito das ações propõe a conversão do recibo.
    Quem já lançou a subscrição do jeito antigo, direto na ação, vê o exercício como lançado."""
    observacao = f"proposto da movimentação da B3 ({m.movimentacao})"
    if convertido:
        dados = DadosEvento(
            tipo=TipoEvento.INCORPORACAO, data=m.data, conta=conta, ticker=convertido,
            quantidade=m.quantidade, ticker_destino=codigo, quantidade_destino=m.quantidade,
            natureza=NaturezaEvento.RECIBO_SUBSCRICAO, observacao=observacao, origem="b3",
        )  # fmt: skip
        nota = NOTA_CONVERSAO
        feito = _lancado_do_recibo(s, m.data, lancados, conversao_de=convertido)
    else:
        recibo = RE_RECIBO.match(codigo)
        if recibo is None:  # _mapa_proposta só manda código de recibo
            raise ImportacaoError(f"{codigo}: código de recibo de subscrição inválido")
        direito = recibo["raiz"] + DIREITO_DO_RECIBO[recibo["classe"]]
        acao = recibo["raiz"] + ACAO_DO_RECIBO[recibo["classe"]]
        dados = DadosEvento(
            tipo=TipoEvento.EXERCICIO_DIREITO, data=m.data, conta=conta, ticker=direito,
            quantidade=m.quantidade, ticker_destino=codigo, quantidade_destino=m.quantidade,
            observacao=observacao, origem="b3",
        )  # fmt: skip
        nota = f"{NOTAS_PROPOSTA[TipoEvento.EXERCICIO_DIREITO]}; {NOTA_CREDITO}"
        feito = _lancado_do_recibo(s, m.data, lancados, exercicio=(direito, codigo, acao))
    if m.chave in ignoradas:
        situacao = "ignorada"
    elif feito:
        situacao = "lançado"
    else:
        situacao = "pendente"
    return Proposta(m.chave, m.data, m.movimentacao, codigo, m.quantidade, m.valor, corretora,
                    situacao, dados, nota)  # fmt: skip


BonificacoesDoPacote = dict[str, list[tuple[date, Decimal, str]]]


def _bonificacoes_do_pacote(
    fiscais: ReferenciasFiscais | None, pasta: Path | None
) -> BonificacoesDoPacote:
    """Ticker recebido -> (data ex, custo atribuído por ação, referência) das bonificações do pacote
    público que trazem o custo (os da CVM). Vazio sem o pacote."""
    if fiscais is None or pasta is None:
        return {}
    por_isin = fontes.tickers_por_isin(pasta)
    saida: BonificacoesDoPacote = defaultdict(list)
    for ev in fontes.eventos(pasta):
        if ev.tipo != "BONIFICACAO" or not ev.custo_atribuido:
            continue
        ticker = por_isin.get(ev.isin_emitido or ev.isin)
        if ticker is None:
            continue
        try:
            data_ex = proximo_pregao(date.fromisoformat(ev.data_com), fiscais.calendario)
            custo = Decimal(ev.custo_atribuido)
        except (ReferenciaError, InvalidOperation, ValueError):
            continue
        if custo > 0:
            saida[ticker].append((data_ex, custo, ev.referencia))
    return saida


def _custo_da_bonificacao(
    bonificacoes: BonificacoesDoPacote, ticker: str, credito: date
) -> tuple[Decimal, str] | None:
    """Custo atribuído por ação da bonificação do pacote no mesmo ativo cuja data ex fica a até
    JANELA dias do crédito que a B3 mostra. Dois custos diferentes na janela: nenhum (o usuário
    confere)."""
    perto = timedelta(days=JANELA)
    achados = {(c, r) for dia, c, r in bonificacoes.get(ticker, []) if abs(credito - dia) <= perto}
    if len({c for c, _ in achados}) != 1:
        return None
    return min(achados)


def bonificacoes_sem_custo_de(lancados: Sequence[EventoLancado]) -> list[EventoLancado]:
    """Bonificações lançadas sem custo atribuído (valor vazio ou zero): o preço médio da venda sai
    errado. Fica de fora o ajuste de quantidade de série de opção, que não tem custo."""
    return [e for e in lancados
            if e.dados.tipo is TipoEvento.BONIFICACAO and not e.dados.valor
            and not RE_SERIE_OPCAO.match(e.dados.ticker)]  # fmt: skip


def bonificacoes_sem_custo(banco: Path, titular: int) -> list[EventoLancado]:
    """As bonificações sem custo atribuído gravadas no banco do titular (sino e tela Eventos)."""
    return bonificacoes_sem_custo_de(listar(banco, titular))


def aviso_sem_custo(e: EventoLancado) -> str:
    """Texto do aviso da tela Eventos para uma bonificação lançada sem custo atribuído."""
    q = e.dados.quantidade
    acoes = f", {q.normalize():f} ações" if q else ""
    return (f"Bonificação de {e.dados.ticker} em {e.dados.data:%d/%m/%Y} ({e.conta_nome}{acoes}) "
            f"lançada sem custo atribuído: {AVISO_SEM_CUSTO}. Sem ele, o preço médio da venda "
            "fica errado.")  # fmt: skip


FOLGA_COBERTURA = timedelta(days=45)  # movimentação que começa ou termina perto das notas basta


SO_NA_MOVIMENTACAO = ("subscrições, restituições de capital, ofertas públicas e eventos que o "
                      "pacote público não tem")  # fmt: skip


def lacuna_movimentacao(banco: Path, titular: int) -> str | None:
    """Aviso quando a movimentação da B3 importada não cobre o período das notas: só ela traz os
    eventos de cada conta, e sem ela as propostas dependem do pacote público. None se cobre."""
    with conectar(banco) as fabrica, fabrica() as s:
        notas = s.execute(select(func.min(NotaCorretagem.data_pregao),
                                 func.max(NotaCorretagem.data_pregao))
                          .join(Conta).where(Conta.titular_id == titular)).one()  # fmt: skip
        movs = s.execute(select(func.min(MovimentacaoB3.data), func.max(MovimentacaoB3.data))
                         .where(MovimentacaoB3.titular_id == titular)).one()  # fmt: skip
    primeira, ultima = notas
    if primeira is None or ultima is None:
        return None
    if movs[0] is None or movs[1] is None:
        return (
            f"Sem a movimentação da B3, {SO_NA_MOVIMENTACAO} não aparecem aqui. Importe a "
            "planilha de movimentação da Área do Investidor em Importar > Planilhas da B3."
        )
    faltas = []
    if movs[0] - primeira > FOLGA_COBERTURA:
        faltas.append(f"começa em {movs[0]:%m/%Y} e as notas em {primeira:%m/%Y}")
    if ultima - movs[1] > FOLGA_COBERTURA:
        faltas.append(f"termina em {movs[1]:%m/%Y} e as notas vão até {ultima:%m/%Y}")
    if not faltas:
        return None
    return (
        f"A movimentação da B3 importada {' e '.join(faltas)}: fora dela, "
        f"{SO_NA_MOVIMENTACAO} não aparecem aqui. Importe a movimentação do período que "
        "falta, se a Área do Investidor tiver."
    )


def _mapa_proposta(
    m: MovimentacaoB3, codigo: str, fiscais: ReferenciasFiscais | None
) -> tuple[TipoEvento, str | None] | None:
    """Tipo e C/V do evento que a linha da movimentação sugere (None: nenhum). A baixa do direito
    não exercido é uma saída; a amortização reduz o custo nos fundos de
    ``ParametrosFiscais.amortizacao_reduz_custo`` (FI-Infra, FIP-IE, FIP comum, FII e Fiagro, e ETF
    de ações desde 2024; pareceres do especialista de 09/11, 13/11 e 24/11/2026)."""
    movimentacao = m.movimentacao.strip().lower()
    if any(x in movimentacao for x in NAO_EXERCIDO):
        return None if m.entrada else (TipoEvento.DIREITO_NAO_EXERCIDO, None)
    mapa = PROPOSTAS.get(movimentacao)
    if mapa is None or (
        mapa[0] in (TipoEvento.SUBSCRICAO, TipoEvento.DIREITO_RECEBIDO) and not m.entrada
    ):
        return None
    if movimentacao == "recibo de subscrição" and RE_RECIBO.match(codigo):
        return TipoEvento.EXERCICIO_DIREITO, None  # o direito sai e o recibo entra com o custo
    if movimentacao == "amortização" and (
        not m.entrada or fiscais is None
        or not fiscais.parametros.amortizacao_reduz_custo(fiscais.tipos.tipo_ou_sem(codigo),
                                                          m.data)
    ):  # fmt: skip
        return None
    return mapa


JANELA_PRAZO_DIREITO = 60  # dias entre o fim do prazo de exercício e a baixa na custódia da B3


def _fim_do_prazo(
    codigo: str, baixa: date, fiscais: ReferenciasFiscais | None, pasta: Path | None
) -> tuple[date, str]:  # fmt: skip
    """Data do direito não exercido: o fim do prazo de exercício, que é decadencial (Lei 6.404
    art. 171 §4º), pelo pacote da B3 (a data de subscrição da companhia; sem ela, o fim do período
    de negociação do direito), nunca depois da baixa; senão, a data da baixa na B3. A data mais
    tarde nunca antecipa a perda (parâmetro direito_vencido_data; parecer do especialista de
    09/11/2026)."""
    sem = (baixa, "data da baixa na B3: se souber o fim do prazo de exercício (aviso aos "
                  "acionistas), use-o")  # fmt: skip
    if fiscais is None or pasta is None:
        return sem
    if fiscais.parametros.texto("direito_vencido_data", baixa) != "fim_do_prazo":
        return (baixa, "data da baixa na B3 (parâmetro direito_vencido_data)")
    fins: list[date] = []
    for ev in fontes.eventos(pasta):
        if ev.tipo != "SUBSCRICAO" or ev.emissor != codigo[:4]:
            continue
        datas = [ev.data_subscricao] if ev.data_subscricao else []
        if " a " in ev.periodo:
            datas.append(ev.periodo.split(" a ")[1].strip())
        try:
            fim = max(date.fromisoformat(x) if "-" in x else datetime.strptime(x, "%d/%m/%Y").date()
                      for x in datas)  # fmt: skip
        except ValueError:
            continue
        # a subscrição que terminou até JANELA_PRAZO_DIREITO dias antes da baixa (a liquidação
        # pode cair uns dias depois da baixa: a data fica na da baixa)
        if timedelta(days=-JANELA) <= baixa - fim <= timedelta(days=JANELA_PRAZO_DIREITO):
            fins.append(min(fim, baixa))
    if not fins:
        return sem
    fim = max(fins)
    return fim, (f"data: fim do prazo de subscrição pelo pacote da B3 ({fim:%d/%m/%Y}); a B3 "
                 f"baixou em {baixa:%d/%m/%Y}")  # fmt: skip


JANELA_RESGATE = 31  # dias entre a data ex e o pagamento em dinheiro que a B3 mostra


AVISO_ANTERIOR = "lance antes o evento anterior deste ativo: a quantidade pode mudar"


def _sem_pendencia_antes(lista: list[Proposta]) -> list[Proposta]:
    """A quantidade de uma proposta é calculada sem os eventos pendentes (100 ações, bonificação de
    10% e depois desdobramento de 100%: o desdobramento são 110, não 100). Por isso só a mais antiga
    de cada ativo é segura, e o número que o usuário confere é o que entra; lançada ela, a seguinte
    passa a ser a mais antiga, já com a quantidade certa. Empate na data também pede conferência."""
    pendentes = [p for p in lista if p.situacao == "pendente"]
    saida: list[Proposta] = []
    for p in lista:
        if p.situacao == "pendente":
            mesmas = [o for o in pendentes if o is not p and o.codigo == p.codigo]
            if any(o.data < p.data for o in mesmas):
                p = replace(p, segura=False, nota=f"{p.nota}; {AVISO_ANTERIOR}")
            elif any(o.data == p.data and (not o.segura or o.dados.conta == p.dados.conta)
                     for o in mesmas):  # fmt: skip
                p = replace(p, segura=False)
        saida.append(p)
    return saida
