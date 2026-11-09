"""Livro de operações -> movimentos de posição, dia a dia (determinístico, sem relógio).

Ordem dentro do dia (por titular): eventos; liquidação de termo; day trade (casado pela ordem dos
negócios, art. 65 §3º; o excedente vira comum pelo preço do próprio negócio); compras;
exercícios; vendas; vencimento (pó).

Fontes:
- Day trade: IN RFB 1.585/2015 art. 65 (mesmo ativo, mesmo dia, mesma corretora desde a
  MP 497/2010; antes, qualquer corretora); §13: exercício de opção não é day trade.
- Exercício (art. 58 §§): titular de call soma o prêmio ao custo; lançador de call soma o prêmio à
  alienação; titular de put subtrai o prêmio da alienação; lançador de put subtrai do custo.
- Opção não exercida: o prêmio é perda (titular) ou ganho (lançador) no vencimento.
- Termo: as ações entram na carteira na liquidação, pelo preço a termo + custos da nota. Vendas
  à vista na mesma data consomem primeiro o lote do termo, pelo preço a termo (art. 62 I, linha do
  termo, sem isenção); o excedente sai da posição anterior pelo PM. O que sobra do lote fica no
  custo médio (art. 62 §1º e art. 58). Decisões do usuário de 20 e 21/10/2026.
"""

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
    Grupo,
    Movimento,
    NaturezaEvento,
    NotaMotor,
    OperacaoMotor,
    PosicaoAtivo,
    TipoAtivo,
    TipoEvento,
    Visao,
)
from coir.domain.parametros import ParametrosFiscais
from coir.engine.posicao import Baixa, Posicao
from coir.engine.rateio import ratear
from coir.erros import MotorError

MERCADOS_DT = frozenset({"VISTA", "FRACIONARIO", "CALL", "PUT"})
MERCADOS_ACAO_A_VISTA = frozenset({"VISTA", "FRACIONARIO", "EXERC CALL", "EXERC PUT"})
MERCADOS_EXERCICIO = frozenset({"EXERC CALL", "EXERC PUT"})
ABREM_POSICAO = frozenset(
    {TipoEvento.SUBSCRICAO, TipoEvento.OFERTA_PUBLICA, TipoEvento.CUSTODIA_INICIAL}
)

# titular, grupo da visão (None por titular; corretora; conta), ativo, vencimento (opção: o
# código da série se repete de um ano para outro, ex.: TAEEL382 em dez/2023 e dez/2025)
Chave = tuple[int, int | str | None, int, date | None]
Origem = tuple[int, str]  # conta, corretora
OrigemMovimento = tuple[int | None, str | None]  # None quando o passo junta contas/corretoras


@dataclass(frozen=True)
class _Op:
    nota: NotaMotor
    op: OperacaoMotor
    custos: Decimal

    @property
    def valor(self) -> Decimal:
        """Custo (compra) ou alienação líquida (venda)."""
        return self.op.total + self.custos if self.op.cv == "C" else self.op.total - self.custos


@dataclass(frozen=True)
class _Peca:
    """Parte de uma operação aplicada à posição (operação inteira ou excedente do day trade)."""

    base: _Op
    quantidade: Decimal
    valor: Decimal
    custos: Decimal
    bruto: Decimal
    regra: str


@dataclass
class _Dia:
    eventos: list[Evento] = field(default_factory=lambda: list[Evento]())
    termos: list[_Op] = field(default_factory=lambda: list[_Op]())
    operacoes: list[_Op] = field(default_factory=lambda: list[_Op]())


def _grupo(tipo: TipoAtivo) -> Grupo:
    return Grupo.OPCOES if tipo is TipoAtivo.OPCAO else Grupo.A_VISTA


def _grupo_opcao(op: OperacaoMotor, params: ParametrosFiscais, dia: date) -> Grupo:
    """Linha da ficha do IRPF para a opção: "opções - ações" se o ativo-objeto é de um dos tipos
    do parâmetro (ação e unit); ETF, índice e outros vão para "opções - outros". Objeto não
    identificado fica em ações, como antes da F8."""
    acoes = {t.strip() for t in params.texto("opcoes_linha_acoes_tipos_objeto", dia).split(";")}
    if op.objeto_tipo is None or op.objeto_tipo.value in acoes:
        return Grupo.OPCOES
    return Grupo.OPCOES_OUTROS


def _categoria(tipo: TipoAtivo, day_trade: bool) -> Categoria:
    if tipo is TipoAtivo.FII:
        return Categoria.FII  # FII (inclusive day trade): 20%, compensação só entre FII
    return Categoria.DAY_TRADE if day_trade else Categoria.COMUM


def _proporcao(valor: Decimal, parte: Decimal, todo: Decimal) -> Decimal:
    return valor if parte == todo else valor * parte / todo


class Livro:
    """Posições de todos os titulares numa visão: a fiscal (por titular, todas as corretoras e
    contas juntas) ou uma operacional (por corretora ou por conta). A conta e a corretora de
    origem vão em cada movimento, em qualquer visão."""

    def __init__(self, visao: Visao) -> None:
        self.visao = visao
        self.posicoes: dict[Chave, Posicao] = defaultdict(Posicao)
        # contas que formam cada posição aberta (atribui o pó à conta, se for uma só)
        self.origens: dict[Chave, set[Origem]] = defaultdict(set)
        self.corretoras: dict[int, str] = {}  # conta -> corretora
        self.tickers: dict[int, str] = {}
        self.tipos: dict[int, TipoAtivo] = {}
        self.grupos: dict[int, Grupo] = {}  # opção: linha pelo ativo-objeto
        # lote do termo liquidado no dia em processamento que ainda não foi vendido nesta data
        # (quantidade e custo pelo preço a termo), por posição
        self.termo_hoje: dict[Chave, Posicao] = {}
        self.parametros = ParametrosFiscais(parametros=())  # os do livro, ao percorrer
        self.fins_pme: set[date] = set()  # dias seguintes ao fim de uma isenção de PME
        self.movimentos: list[Movimento] = []
        self.avisos: list[str] = []
        # titular e ativos do item em processamento (o que isolar se o motor parar nele)
        self.em_curso: tuple[int, frozenset[int]] | None = None
        # (titular, ativo) -> dia a partir do qual o ativo ficou fora do cálculo
        self.cortes: dict[tuple[int, int], date] = {}
        # ativo puxado para fora por outro ligado (exercício, incorporação) -> o que o puxou
        self.puxados: dict[tuple[int, int], int] = {}
        # venda que passou de comprado a vendido: posição, dia, quantidade vendida, posição antes
        self.suspeitas: list[tuple[Chave, date, Decimal, Decimal]] = []

    def fora(self, titular: int, ativo: int, dia: date) -> bool:
        corte = self.cortes.get((titular, ativo))
        return corte is not None and corte <= dia

    def processando(self, titular: int, *ativos: int | None) -> None:
        self.em_curso = (titular, frozenset(a for a in ativos if a is not None))
        for ativo in self.em_curso[1]:
            if self.tipos.get(ativo) is TipoAtivo.SEM_TIPO:
                raise MotorError(
                    f"{self.tickers.get(ativo, ativo)} não tem tipo de ativo nas tabelas do COIR, "
                    "e sem o tipo não se sabe como ele é tributado. Atualize as tabelas em "
                    "Configurações; se continuar, mande a nota pela tela Suporte"
                )

    def _grupo(self, de: Origem) -> int | str | None:
        if self.visao is Visao.CONTA:
            return de[0]
        return de[1] if self.visao is Visao.CORRETORA else None

    def chave(self, titular: int, de: Origem, ativo: int, vencimento: date | None = None) -> Chave:
        self.corretoras[de[0]] = de[1]
        return (titular, self._grupo(de), ativo, vencimento)

    def chave_opcao(self, titular: int, de: Origem, ativo: int, dia: date) -> Chave | None:
        """Posição aberta da série com o vencimento mais próximo a partir de `dia`."""
        grupo = self._grupo(de)
        abertas = [
            k for k, p in self.posicoes.items()
            if k[:3] == (titular, grupo, ativo) and p.quantidade
            and (k[3] is None or k[3] >= dia)
        ]  # fmt: skip
        return min(abertas, key=lambda k: k[3] or date.max, default=None)

    def origem_unica(self, chave: Chave) -> OrigemMovimento:
        contas = self.origens.get(chave, set())
        if len(contas) == 1:
            return next(iter(contas))
        corretoras = {c for _, c in contas}
        return (None, next(iter(corretoras)) if len(corretoras) == 1 else None)

    def registrar(
        self, chave: Chave, dia: date, origem: str, regra: str, *, de: OrigemMovimento,
        categoria: Categoria, cv: str | None, quantidade: Decimal, valor: Decimal,
        custos: Decimal = ZERO, baixa: Baixa | None = None, venda_acao: Decimal = ZERO,
        isentavel: bool = False, mes_limite: date | None = None, isento_pme: bool = False,
        toca_posicao: bool = True, base_irrf: Decimal = ZERO, grupo: Grupo | None = None,
        evento: TipoEvento | None = None,
    ) -> None:  # fmt: skip
        titular, _, ativo, _ = chave
        pos = self.posicoes[chave]
        conta, corretora = de
        if toca_posicao and conta is not None and corretora is not None:
            self.origens[chave].add((conta, corretora))
        if pos.quantidade == 0:
            self.origens.pop(chave, None)
        tipo = self.tipos[ativo]
        b = baixa or Baixa()
        self.movimentos.append(
            Movimento(
                data=dia, titular=titular, conta=conta, corretora=corretora, ativo=ativo,
                ticker=self.tickers[ativo],
                origem=origem, regra=regra, categoria=categoria,
                grupo=grupo or self.grupos.get(ativo) or _grupo(tipo), tipo=tipo,
                cv=cv,  # type: ignore[arg-type]
                quantidade=quantidade, valor=valor, custos=custos, custo_baixa=b.custo_baixa,
                resultado=b.resultado, venda_acao_a_vista=venda_acao, isentavel=isentavel,
                mes_limite=mes_limite, isento_pme=isento_pme, base_irrf=base_irrf,
                quantidade_apos=pos.quantidade, custo_apos=pos.custo, evento=evento,
            )
        )  # fmt: skip


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


Meses = dict[tuple[int, int], set[tuple[int, int]]]  # (titular, ativo) -> (ano, mês)
OpsDoDia = list[tuple[date, _Op]]


def _desde_zerado(livro: Livro, meses: Meses) -> None:
    """O evento que falta pode ser anterior ao dia em que o motor parou: os meses do ativo desde
    a última vez que a posição zerou também ficam incompletos (decisão do usuário de 03/11/2026,
    parecer do especialista)."""
    for (titular, ativo), corte in livro.cortes.items():
        movs = [m for m in livro.movimentos
                if m.titular == titular and m.ativo == ativo and m.data < corte]  # fmt: skip
        zerou = max((i for i, m in enumerate(movs) if m.quantidade_apos == 0), default=-1)
        meses[(titular, ativo)] |= {(m.data.year, m.data.month) for m in movs[zerou + 1 :]}


def _venda_isolada(livro: Livro, dia: date, o: _Op) -> None:
    """Venda de ativo isolado: o resultado fica de fora (o custo não é conhecido), mas o valor
    da venda é: entra na soma do limite de isenção do mês e na base que reparte o IRRF (decisão
    do usuário de 03/11/2026). O IRRF da nota continua crédito do mês."""
    op = o.op
    if op.cv != "V" or op.mercado == "TERMO":
        return
    pme = op.pme_ate is not None and dia <= op.pme_ate
    acao = op.tipo in (TipoAtivo.ACAO, TipoAtivo.UNIT) and op.mercado in MERCADOS_ACAO_A_VISTA
    base_irrf = op.total if (op.mercado not in MERCADOS_EXERCICIO and not pme
                             and op.tipo is not TipoAtivo.ETF_RF) else ZERO  # fmt: skip
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


def _isolar(
    livro: Livro, ops: OpsDoDia, eventos: list[Evento]
) -> tuple[OpsDoDia, OpsDoDia, list[Evento], Meses]:
    """Tira do livro as operações e os eventos de ativos isolados a partir do corte. O que está
    ligado vai junto: o exercício leva a série exercida; a incorporação e a cisão, o destino."""
    meses: Meses = defaultdict(set)
    if not livro.cortes:
        return ops, [], eventos, meses
    itens = [(dia, o.nota.titular, (o.op.ativo, o.op.opcao)) for dia, o in ops] + [
        (e.data, e.titular, (e.ativo, e.ativo_destino)) for e in eventos
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
    eventos = [e for e in eventos if not tirar(e.data, e.titular, (e.ativo, e.ativo_destino))]
    return fica, tiradas, eventos, meses


def _percorrer(
    livro: Livro, notas: list[NotaMotor], eventos: list[Evento], parametros: ParametrosFiscais,
    ate: date,
) -> Meses:  # fmt: skip
    """Monta os dias e processa; devolve os meses com itens tirados por ativo isolado."""
    dias: dict[date, _Dia] = defaultdict(_Dia)
    vencimentos: set[date] = set()
    ops: list[tuple[date, _Op]] = []  # dia em que a operação entra no livro
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
            item = _Op(nota, op, custo)
            if op.mercado == "TERMO":
                if op.cv == "V" or op.liquidacao_termo is None:
                    raise MotorError(f"nota {nota.numero} de {nota.data}, {op.ticker}: venda a "
                                     "termo ou termo sem data de liquidação (prazo)")  # fmt: skip
                if op.liquidacao_termo <= ate:
                    ops.append((op.liquidacao_termo, item))
                continue
            ops.append((nota.data, item))
            if op.tipo is TipoAtivo.OPCAO:
                if op.vencimento is None or op.vencimento < nota.data:
                    raise MotorError(
                        f"opção {op.ticker} em {nota.data}: vencimento {op.vencimento} ausente ou "
                        "anterior ao negócio (confira data/vencimentos_opcoes*.csv)"
                    )
                vencimentos.add(op.vencimento)
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
    ops, tiradas, ate_hoje, meses = _isolar(livro, ops, ate_hoje)
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
    for venc in vencimentos:
        if venc <= ate:
            dias.setdefault(venc, _Dia())  # o pó acontece mesmo sem negócio no dia
    ofertas = [e for e in eventos if e.tipo is TipoEvento.OFERTA_PUBLICA
               and not livro.fora(e.titular, e.ativo, e.data)]  # fmt: skip
    for dia in sorted(dias):
        try:
            _antecipar_ofertas(livro, dia, dias, ofertas, parametros)
            _processar_dia(livro, dia, dias[dia], parametros)
        except MotorError as erro:
            erro.data = erro.data or dia
            raise
    livro.em_curso = None
    _avisar_suspeitas(livro)
    for dia, o in tiradas:
        _venda_isolada(livro, dia, o)
    return meses


def _antecipar_ofertas(
    livro: Livro, dia: date, dias: dict[date, _Dia], ofertas: list[Evento],
    parametros: ParametrosFiscais,
) -> None:  # fmt: skip
    """A negociação pode começar antes da liquidação da oferta: a venda do dia que deixaria a
    posição vendida é das ações da oferta, então a aquisição entra antes dela, e o ganho fica no
    mês do pregão da venda (decisão do usuário de 08/11/2026; IN RFB 1.585/2015 art. 58). Antecipa
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


def _processar_dia(livro: Livro, dia: date, d: _Dia, params: ParametrosFiscais) -> None:
    if dia in livro.fins_pme:
        _fim_da_isencao(livro, dia)
    # cisões depois dos demais eventos do dia (agrupadas por posição de origem e destino)
    cisoes: dict[tuple[Chave, int | None], list[Evento]] = defaultdict(list)
    for evento in sorted(d.eventos, key=lambda e: e.id):
        livro.processando(evento.titular, evento.ativo, evento.ativo_destino)
        if evento.tipo is TipoEvento.CISAO:
            # as contas que formam a posição (visão por titular) entram juntas: o custo repartido
            # é o de antes do evento, não o que sobra depois da primeira conta
            cisoes[(_chave_evento(livro, evento), evento.ativo_destino)].append(evento)
        elif evento.tipo is TipoEvento.INCORPORACAO:
            _incorporacao(livro, evento, params)
        elif evento.tipo is TipoEvento.LEILAO_FRACAO:
            _leilao_fracao(livro, evento, params)
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
    for t in sorted(d.termos, key=lambda t: (t.nota.id, t.op.ordem)):
        livro.processando(t.nota.titular, t.op.ativo)
        chave = livro.chave(t.nota.titular, _de(t), t.op.ativo)
        pos = livro.posicoes[chave]
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
    mesma_corretora = params.booleano("day_trade_mesma_corretora", dia)
    pela_marcacao = params.texto("day_trade_criterio", dia) == "marcacao_corretora"
    livro.em_curso = None
    compras, vendas, exercicios = _day_trade(livro, dia, d.operacoes, mesma_corretora,
                                             pela_marcacao)  # fmt: skip
    for p in sorted(compras, key=_ordem):
        livro.processando(p.base.nota.titular, p.base.op.ativo)
        _aplicar(livro, dia, p)
    for o in sorted(exercicios, key=lambda o: (o.nota.id, o.op.ordem)):
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


def _de(o: _Op) -> Origem:
    return (o.nota.conta, o.nota.corretora)


def _ordem(p: _Peca) -> tuple[int, int]:
    return (p.base.nota.id, p.base.op.ordem)


def _day_trade(
    livro: Livro, dia: date, ops: Sequence[_Op], mesma_corretora: bool, pela_marcacao: bool
) -> tuple[list[_Peca], list[_Peca], list[_Op]]:
    """Casa compras e vendas do mesmo ativo no dia; devolve as peças comuns (compras, vendas) e os
    exercícios (nunca day trade). Com `pela_marcacao` (parâmetro day_trade_criterio), os negócios
    que a corretora marcou D formam o day trade quando a marcação é consistente; senão, a ordem
    dos negócios impressa (IN RFB 1.585/2015 art. 65 §3º)."""
    compras: list[_Peca] = []
    vendas: list[_Peca] = []
    exercicios: list[_Op] = []
    grupos: dict[tuple[int, str | None, int, date | None], list[_Op]] = defaultdict(list)
    for o in ops:
        if o.op.mercado in MERCADOS_EXERCICIO:
            exercicios.append(o)
        elif o.op.mercado in MERCADOS_DT:
            corretora = o.nota.corretora if mesma_corretora else None
            grupos[(o.nota.titular, corretora, o.op.ativo, o.op.vencimento)].append(o)
        else:
            raise MotorError(f"mercado {o.op.mercado} sem regra de cálculo")
    for grupo in grupos.values():
        c = sorted((o for o in grupo if o.op.cv == "C"), key=_ordem_negocio)
        v = sorted((o for o in grupo if o.op.cv == "V"), key=_ordem_negocio)
        qc = sum((o.op.quantidade for o in c), ZERO)
        qv = sum((o.op.quantidade for o in v), ZERO)
        dt = min(qc, qv)
        if dt == 0:
            compras += [_peca_inteira(o) for o in c]
            vendas += [_peca_inteira(o) for o in v]
            continue
        regra = ORDEM
        if pela_marcacao and any(o.op.marcada_dt for o in grupo):
            mc = [o for o in c if o.op.marcada_dt]
            mv = [o for o in v if o.op.marcada_dt]
            if (
                sum((o.op.quantidade for o in mc), ZERO) >= dt
                and sum((o.op.quantidade for o in mv), ZERO) >= dt
            ):
                # os negócios marcados D primeiro: cobrem o day trade do dia, e a corretora pode
                # marcar a linha inteira quando só parte dela fecha no dia (500 D x 200 D: 200 de
                # day trade e 300 comuns, IN RFB 1.585/2015 art. 65 §§1º a 3º)
                c = mc + [o for o in c if not o.op.marcada_dt]
                v = mv + [o for o in v if not o.op.marcada_dt]
                regra = MARCACAO
            else:
                livro.avisos.append(
                    f"{dia}: day trade de {grupo[0].op.ticker}: a marcação D da nota não fecha "
                    f"com a quantidade do dia ({dt}); casado pela ordem dos negócios impressa"
                )
        dt_c, sobra_c = _fatiar(c, dt, "compra")
        dt_v, sobra_v = _fatiar(v, dt, "venda")
        _movimento_day_trade(livro, dia, grupo, dt, dt_c, dt_v, regra)
        compras += sobra_c
        vendas += sobra_v
    return compras, vendas, exercicios


ORDEM = "day trade (IN 1.585 art. 65 §3º: 1ª compra com a 1ª venda, pela ordem impressa)"
MARCACAO = "day trade (negócios marcados D pela corretora, que tem o horário; IN 1.585 art. 65 §3º)"


def _ordem_negocio(o: _Op) -> tuple[date, str, int, int]:
    """Ordem dos negócios no dia: número da nota e ordem impressa (a nota não traz horário)."""
    numero = o.nota.numero
    return (o.nota.data, numero.zfill(20) if numero.isdigit() else numero, o.op.ordem, o.op.id)


def _peca_inteira(o: _Op) -> _Peca:
    regra = "compra: custo = valor + custos" if o.op.cv == "C" else (
        "venda: resultado = valor - custos - PM x quantidade")  # fmt: skip
    return _Peca(o, o.op.quantidade, o.valor, o.custos, o.op.total, regra)


def _fatiar(ops: list[_Op], dt: Decimal, lado: str) -> tuple[list[_Peca], list[_Peca]]:
    """IN RFB 1.585/2015 art. 65 §3º: casa pela ordem (1º negócio de compra com o 1º de venda).
    Os primeiros negócios até a quantidade de day trade formam o day trade; o que sobra (inclusive
    a parte final de um negócio) é operação comum pelo preço do próprio negócio."""
    casadas: list[_Peca] = []
    sobra: list[_Peca] = []
    falta = dt
    for o in ops:
        q = o.op.quantidade
        usa = min(falta, q)
        falta -= usa
        if usa:
            casadas.append(_parte(o, usa, "day trade"))
        if usa < q:
            sobra.append(_parte(o, q - usa, f"excedente do day trade ({lado}): operação comum"))
    return casadas, sobra


def _parte(o: _Op, quantidade: Decimal, regra: str) -> _Peca:
    q = o.op.quantidade
    return _Peca(o, quantidade, _proporcao(o.valor, quantidade, q),
                 _proporcao(o.custos, quantidade, q), _proporcao(o.op.total, quantidade, q),
                 regra)  # fmt: skip


def _movimento_day_trade(
    livro: Livro, dia: date, grupo: list[_Op], dt: Decimal, compras: list[_Peca],
    vendas: list[_Peca], regra: str,
) -> None:  # fmt: skip
    primeiro = min(grupo, key=_de)
    contas = {_de(o) for o in grupo}
    # negócios em contas diferentes (da mesma corretora; antes de 28/07/2010, de qualquer uma): o
    # movimento fica sem conta (e sem corretora, se forem várias); a origem lista as operações
    corretoras = {c for _, c in contas}
    de: OrigemMovimento = _de(primeiro) if len(contas) == 1 else (
        None, primeiro.nota.corretora if len(corretoras) == 1 else None)  # fmt: skip
    chave = livro.chave(primeiro.nota.titular, _de(primeiro), primeiro.op.ativo,
                        primeiro.op.vencimento)  # fmt: skip
    custo = sum((p.valor for p in compras), ZERO)
    alienacao = sum((p.valor for p in vendas), ZERO)
    custos = sum((p.custos for p in (*compras, *vendas)), ZERO)
    origem = "op:" + ",".join(str(o.op.id) for o in sorted(grupo, key=lambda o: o.op.id))
    tipo = primeiro.op.tipo
    # day trade com ações de PME no período: isento (leitura literal do art. 59 §2º I, que só
    # afasta o limite de R$ 20 mil; decisão do usuário de 15/10/2026)
    pme = all(p.base.op.pme_ate is not None and dia <= p.base.op.pme_ate for p in compras)
    livro.registrar(chave, dia, origem,
                    f"{regra}: resultado = vendas casadas - compras casadas", de=de,
                    toca_posicao=False,
                    categoria=_categoria(tipo, day_trade=True), cv=None, quantidade=dt,
                    valor=alienacao, custos=custos, isento_pme=pme,
                    baixa=Baixa(resultado=alienacao - custo, custo_baixa=custo))  # fmt: skip


def _aplicar(livro: Livro, dia: date, p: _Peca) -> None:
    o = p.base
    chave = livro.chave(o.nota.titular, _de(o), o.op.ativo, o.op.vencimento)
    lote = livro.termo_hoje.get(chave)
    if o.op.cv != "V" or o.op.mercado in MERCADOS_EXERCICIO or lote is None or not lote.quantidade:
        _aplicar_parte(livro, dia, chave, p)
        return
    # venda na data da liquidação do termo (decisão do usuário de 21/10/2026, parecer do
    # especialista): consome primeiro o lote do termo, pela ordem dos negócios; o excedente sai da
    # posição anterior pelo PM, como venda à vista comum
    q = min(p.quantidade, lote.quantidade)
    custo = lote.baixar_sem_resultado(q)
    regra = (p.regra + "; venda na data da liquidação do termo: custo = preço a termo + custos da "
             "compra (IN 1.585 art. 62 I)")  # fmt: skip
    _aplicar_parte(livro, dia, chave, _fatia(p, q, regra), custo_termo=custo)
    if q < p.quantidade:
        regra = (p.regra + "; venda na data da liquidação do termo, além do lote do termo: sai da "
                 "posição anterior pelo PM")  # fmt: skip
        _aplicar_parte(livro, dia, chave, _fatia(p, p.quantidade - q, regra))


def _fatia(p: _Peca, quantidade: Decimal, regra: str) -> _Peca:
    return _Peca(p.base, quantidade, _proporcao(p.valor, quantidade, p.quantidade),
                 _proporcao(p.custos, quantidade, p.quantidade),
                 _proporcao(p.bruto, quantidade, p.quantidade), regra)  # fmt: skip


def _vender_por_lote(livro: Livro, dia: date, chave: Chave, p: _Peca) -> None:
    """IN 1.585 art. 69-A §4º: cada lote tem o seu PM. A norma não fixa qual sai primeiro
    (parâmetro pme_lote_ordem_venda; decisão do usuário de 08/11/2026). O que passar da posição é
    venda a descoberto, como em qualquer ativo; a venda inteira se reparte pela quantidade."""
    pos = livro.posicoes[chave]
    antes = pos.quantidade
    lotes = [("isento", pos.isento_quantidade),
             ("tributado", pos.quantidade - pos.isento_quantidade)]  # fmt: skip
    if livro.parametros.texto("pme_lote_ordem_venda", dia) == "tributado_primeiro":
        lotes.reverse()
    falta = p.quantidade
    for lote, tem in lotes:
        q = min(falta, tem)
        if q:
            regra = f"{p.regra}; lote {lote} de PME (IN 1.585 art. 69-A §4º)"
            _aplicar_parte(livro, dia, chave, _fatia(p, q, regra), lote=lote)
            falta -= q
    if falta:
        _aplicar_parte(livro, dia, chave, _fatia(p, falta, p.regra))
        if livro.visao is Visao.TITULAR and p.base.op.mercado not in MERCADOS_EXERCICIO:
            livro.suspeitas.append((chave, dia, p.quantidade, antes))


def _fim_da_isencao(livro: Livro, dia: date) -> None:
    """Lei 13.043/2014 art. 17 §6º e IN 1.585 art. 68: no dia seguinte ao fim da isenção, o lote
    isento sobe à média dos 30 pregões, se maior que o custo pago, e funde-se ao tributado (art.
    58). A posição de 31/12/2023 continua pelo custo pago."""
    for chave, pos in list(livro.posicoes.items()):
        if pos.isento_quantidade and pos.pme_ate is not None and pos.pme_ate < dia:
            titular, _, ativo, _ = chave
            livro.processando(titular, ativo)
            aumento = pos.fim_da_isencao()
            if aumento:
                livro.registrar(
                    chave, dia, "pme:fim",
                    "fim da isenção de PME: custo do lote isento ajustado à média dos 30 pregões "
                    "(Lei 13.043/2014 art. 17 §6º)",
                    de=(None, None), toca_posicao=False, categoria=Categoria.COMUM, cv=None,
                    quantidade=ZERO, valor=aumento,
                )  # fmt: skip


def _aplicar_parte(
    livro: Livro, dia: date, chave: Chave, p: _Peca, custo_termo: Decimal | None = None,
    lote: str | None = None,
) -> None:  # fmt: skip
    """Aplica a peça à posição. Com `custo_termo`, é a venda de ações do lote do termo na data da
    liquidação: sai pelo preço a termo, vai para a linha "Mercado a termo" da ficha do IRPF e nunca
    é isenta (IN 1.585 art. 62 I e art. 59 §2º IV; decisão do usuário de 20/10/2026)."""
    o = p.base
    pos = livro.posicoes[chave]
    if (o.op.cv == "V" and custo_termo is None and lote is None and pos.quantidade > 0
            and pos.isento_em(dia)):  # fmt: skip
        _vender_por_lote(livro, dia, chave, p)
        return
    # ação ou unit no mercado à vista (units só entram com o parâmetro units_na_isencao)
    acao = o.op.tipo in (TipoAtivo.ACAO, TipoAtivo.UNIT) and o.op.mercado in MERCADOS_ACAO_A_VISTA
    # IN 1.585 art. 59 §2º IV: a isenção não se aplica à alienação por exercício de opção; o valor
    # continua na soma do limite (decisão do usuário de 15/10/2026, leitura conservadora)
    isentavel = acao and o.op.mercado not in MERCADOS_EXERCICIO
    # venda a descoberto (leitura A1, decisão do usuário de 15/10/2026, sem norma específica): o
    # resultado sai na recompra, mas o limite de isenção é o do mês da venda
    mes_limite = pos.descoberto_desde if o.op.cv == "C" and pos.quantidade < 0 else None
    pme = False
    if o.op.cv == "C":
        antes = pos.quantidade
        baixa = pos.comprar(p.quantidade, p.valor)
        _lote_pme(pos, o.op, dia, p.quantidade, p.valor, antes)
        venda_acao = ZERO
        if pos.quantidade >= 0:
            pos.descoberto_desde = None
    else:
        ja_vendida = pos.quantidade < 0
        antes = pos.quantidade
        # Lei 13.043/2014 art. 16 (IN 1.585 arts. 66 a 69-A): ganho isento, fora do limite
        pme = lote == "isento"
        if pme:
            baixa = pos.vender_isento(p.quantidade, p.valor)
        elif lote == "tributado":
            baixa = pos.vender_tributado(p.quantidade, p.valor)
        elif custo_termo is None:
            baixa = pos.vender(p.quantidade, p.valor)
        elif pos.quantidade >= p.quantidade:
            baixa = pos.vender_lote(p.quantidade, p.valor, custo_termo)
        else:
            raise MotorError(f"{dia}: venda de {o.op.ticker} na data da liquidação do termo maior "
                             f"que a posição ({pos.quantidade})")  # fmt: skip
        venda_acao = p.bruto if acao and not pme else ZERO
        if pos.quantidade < 0 and not ja_vendida:
            pos.descoberto_desde = dia
            if (
                antes > 0
                and o.op.tipo not in (TipoAtivo.OPCAO, TipoAtivo.FUTURO)
                and (o.op.mercado not in MERCADOS_EXERCICIO and livro.visao is Visao.TITULAR)
            ):
                livro.suspeitas.append((chave, dia, p.quantidade, antes))
    # venda a descoberto a partir de posição zerada é operação normal, sem aviso (decisão do
    # usuário de 25/10/2026); passar de comprado a vendido e continuar vendido no fim do mês
    # costuma ser evento que falta: só avisa, o cálculo não muda (decisão de 03/11/2026)
    # base do IRRF 0,005% (Lei 11.033/2004 art. 2º §§1º e 2º): venda comum, sem exercício, PME
    # nem ETF de renda fixa (tributação exclusiva); usada só para repartir o IRRF entre as fichas
    base_irrf = p.bruto if (o.op.cv == "V" and o.op.mercado not in MERCADOS_EXERCICIO and not pme
                            and o.op.tipo is not TipoAtivo.ETF_RF) else ZERO  # fmt: skip
    termo = custo_termo is not None
    livro.registrar(chave, dia, f"op:{o.op.id}", p.regra, de=_de(o),
                    categoria=_categoria(o.op.tipo, day_trade=False), cv=o.op.cv,
                    quantidade=p.quantidade, valor=p.valor, custos=p.custos, baixa=baixa,
                    venda_acao=venda_acao, isentavel=isentavel and not pme and not termo,
                    mes_limite=mes_limite, isento_pme=pme, base_irrf=base_irrf,
                    grupo=Grupo.TERMO if termo else None)  # fmt: skip


def _lote_pme(pos: Posicao, op: OperacaoMotor, dia: date, quantidade: Decimal, valor: Decimal,
              antes: Decimal) -> None:  # fmt: skip
    """IN 1.585 art. 69-A §4º: a compra elegível vai para o lote isento; a outra, para o tributado.
    O que cobriu venda a descoberto (`antes` < 0) não fica na posição."""
    sobra = quantidade - min(quantidade, max(-antes, ZERO))
    if sobra and op.pme_ate is not None and dia <= op.pme_ate:
        pos.somar_isento(sobra, valor * sobra / quantidade, op.pme_ate, op.pme_media_fim)


REGRA_EXERCICIO = {
    ("EXERC CALL", "C"): "exercício: titular de call compra; o prêmio pago soma ao custo",
    ("EXERC CALL", "V"): "exercício: lançador de call vende; o prêmio recebido soma à alienação",
    ("EXERC PUT", "V"): "exercício: titular de put vende; o prêmio pago reduz a alienação",
    ("EXERC PUT", "C"): "exercício: lançador de put compra; o prêmio recebido reduz o custo",
}


def _exercicio(livro: Livro, dia: date, o: _Op) -> None:
    op = o.op
    if op.opcao is None:
        raise MotorError(f"exercício {op.ticker} ({dia}) sem a opção exercida")
    chave_opcao = livro.chave_opcao(o.nota.titular, _de(o), op.opcao, dia)
    if chave_opcao is None:
        chave_opcao = livro.chave(o.nota.titular, _de(o), op.opcao)
    opcao = livro.posicoes[chave_opcao]
    titular = (op.mercado, op.cv) in (("EXERC CALL", "C"), ("EXERC PUT", "V"))
    if (opcao.quantidade > 0) != titular or abs(opcao.quantidade) < op.quantidade:
        lado = "comprada" if titular else "vendida"
        raise MotorError(
            f"exercício de {op.quantidade} {livro.tickers.get(op.opcao, op.opcao)} em {dia}: "
            f"posição {lado} na opção é {opcao.quantidade}"
        )
    premio = opcao.baixar_sem_resultado(op.quantidade)  # pago (> 0) ou recebido (< 0)
    livro.registrar(chave_opcao, dia, f"op:{op.id}", "opção exercida: o prêmio vai para o "
                    "ativo-objeto (sem resultado na opção)", de=_de(o), categoria=Categoria.COMUM,
                    cv=None, quantidade=op.quantidade, valor=premio)  # fmt: skip
    valor = o.valor + premio if op.cv == "C" else o.valor - premio
    _aplicar(livro, dia, _Peca(o, op.quantidade, valor, o.custos, op.total,
                               REGRA_EXERCICIO[(op.mercado, op.cv)]))  # fmt: skip


def _chave_evento(livro: Livro, e: Evento) -> Chave:
    de = (e.conta, e.corretora)
    chave = livro.chave(e.titular, de, e.ativo)
    if e.tipo_ativo is TipoAtivo.OPCAO:  # evento na série (ajuste da B3): a posição aberta
        chave = livro.chave_opcao(e.titular, de, e.ativo, e.data) or chave
    return chave


def _nome(e: Evento) -> str:
    return f"{e.tipo.value.lower()} de {e.ticker} ({e.data})"


def _origem_do_evento(livro: Livro, e: Evento, chave: Chave) -> Posicao:
    """Posição de A (origem de incorporação, cisão ou leilão de fração): tem de existir, ter a
    quantidade do evento e não ser lote de PME nem venda a descoberto de ação."""
    pos = livro.posicoes[chave]
    if not e.quantidade or e.quantidade <= 0:
        raise MotorError(f"{_nome(e)} sem a quantidade de origem")
    if pos.quantidade == 0 or e.quantidade > abs(pos.quantidade):
        raise MotorError(f"{_nome(e)}: quantidade {e.quantidade} e posição {pos.quantidade} na "
                         f"conta {e.corretora} {e.conta} (confira data, conta e "
                         "quantidade)")  # fmt: skip
    if pos.quantidade < 0 and e.tipo_ativo is not TipoAtivo.OPCAO:
        raise MotorError(f"{_nome(e)}: posição vendida a descoberto na data do evento")
    if pos.isento_em(e.data):  # sem norma para a isenção na ação recebida e sem caso real
        raise MotorError(f"{_nome(e)}: lote isento de PME dentro do prazo da isenção")
    return pos


def _destino(livro: Livro, e: Evento, chave: Chave, sinal: int) -> Chave:
    """Chave de B (a série de destino herda o vencimento da de origem). Posição de B no sentido
    contrário (ex.: B vendido antes da data do evento) falha alto: o evento cobriria a venda."""
    if e.ativo_destino is None:
        raise MotorError(f"{_nome(e)} sem o ativo de destino")
    destino = livro.chave(e.titular, (e.conta, e.corretora), e.ativo_destino, chave[3])
    if livro.posicoes[destino].quantidade * sinal < 0:
        raise MotorError(
            f"{_nome(e)}: posição em {e.ticker_destino} no sentido contrário "
            f"({livro.posicoes[destino].quantidade}); venda de B antes do evento? (use a data ex)"
        )
    return destino


def _incorporacao(livro: Livro, e: Evento, params: ParametrosFiscais) -> None:
    """Incorporação, fusão, conversão ou ajuste de série: a quantidade de A sai sem resultado e
    B entra com o mesmo custo (IN RFB 1.585/2015 art. 58 §6º; ajuste de série da B3: sem efeito
    fiscal, o custo é o dos prêmios, art. 60 §2º). A data de aquisição continua a original."""
    if e.natureza is None or not e.quantidade_destino or e.quantidade_destino <= 0:
        raise MotorError(f"{_nome(e)} sem a natureza ou a quantidade de destino")
    if e.natureza is NaturezaEvento.INCORPORACAO_ACOES:
        tratamento = params.texto("incorporacao_acoes_tratamento", e.data)
        if tratamento != "custo_mantido":
            raise MotorError(f"incorporação de ações: tratamento {tratamento!r} não implementado")
    chave = _chave_evento(livro, e)
    pos = _origem_do_evento(livro, e, chave)
    sinal = 1 if pos.quantidade > 0 else -1
    destino = _destino(livro, e, chave, sinal)
    de = (e.conta, e.corretora)
    quantidade = e.quantidade or ZERO
    custo = pos.baixar_sem_resultado(quantidade)
    regra = f"incorporação ({e.natureza.value.lower()}, IN 1.585 art. 58 §6º): "
    livro.registrar(chave, e.data, f"evento:{e.id}",
                    regra + f"sai para {e.ticker_destino} com o custo, sem resultado", de=de,
                    categoria=Categoria.COMUM, cv="V" if sinal > 0 else "C",
                    quantidade=quantidade, valor=custo)  # fmt: skip
    livro.posicoes[destino].ajustar(e.quantidade_destino * sinal, custo)
    livro.registrar(destino, e.data, f"evento:{e.id}",
                    regra + f"entra no lugar de {e.ticker} com o mesmo custo", de=de,
                    categoria=Categoria.COMUM, cv="C" if sinal > 0 else "V",
                    quantidade=e.quantidade_destino, valor=custo)  # fmt: skip


def _cisao(livro: Livro, chave: Chave, grupo: list[Evento]) -> None:
    """Cisão: o percentual do custo de A divulgado pela companhia vai para B, sem resultado (IN
    RFB 1.585/2015 art. 58 §6º: na proporção fixada pela assembleia). A quantidade de A não
    muda. Os eventos das contas que formam a mesma posição (visão por titular) somam a quantidade
    de origem, que tem de ser a posição inteira."""
    e0 = grupo[0]
    percentuais = {e.percentual_custo for e in grupo}
    if None in percentuais:
        raise MotorError(f"{_nome(e0)} sem o percentual de custo: informe o da companhia (aviso "
                         "aos acionistas) no evento")  # fmt: skip
    if len(percentuais) > 1:
        raise MotorError(f"{_nome(e0)}: percentuais de custo diferentes entre as contas")
    p = e0.percentual_custo or ZERO
    if not ZERO < p < 1:
        raise MotorError(f"{_nome(e0)}: percentual de custo {p} fora do intervalo (0, 1)")
    if any(not e.quantidade_destino or e.quantidade_destino <= 0 for e in grupo):
        raise MotorError(f"{_nome(e0)} sem a quantidade de {e0.ticker_destino} recebida")
    pos = livro.posicoes[chave]
    for e in grupo:
        _origem_do_evento(livro, e, chave)
    base = sum((e.quantidade or ZERO for e in grupo), ZERO)
    if base != pos.quantidade:
        raise MotorError(f"{_nome(e0)}: quantidade de origem {base} diferente da posição "
                         f"{pos.quantidade} (lance o evento de cada conta)")  # fmt: skip
    total = pos.custo * p  # custo de antes do evento
    destino = _destino(livro, e0, chave, 1)
    restante = total
    for i, e in enumerate(grupo):
        custo = restante if i == len(grupo) - 1 else total * (e.quantidade or ZERO) / base
        restante -= custo
        de = (e.conta, e.corretora)
        pos.ajustar(ZERO, -custo)
        livro.registrar(chave, e.data, f"evento:{e.id}",
                        f"cisão: {p * 100}% do custo vai para {e.ticker_destino}, sem resultado "
                        "(a quantidade não muda)", de=de, categoria=Categoria.COMUM, cv="V",
                        quantidade=ZERO, valor=custo)  # fmt: skip
        livro.posicoes[destino].ajustar(e.quantidade_destino or ZERO, custo)
        livro.registrar(destino, e.data, f"evento:{e.id}",
                        f"cisão: recebida de {e.ticker} com {p * 100}% do custo", de=de,
                        categoria=Categoria.COMUM, cv="C",
                        quantidade=e.quantidade_destino or ZERO, valor=custo)  # fmt: skip


def _leilao_fracao(livro: Livro, e: Evento, params: ParametrosFiscais) -> None:
    """Fração vendida em leilão pela companhia: alienação pelo valor líquido recebido, custo pelo
    PM (IN RFB 1.585/2015 art. 58 caput). Sem norma sobre a isenção: parâmetro
    leilao_fracao_isentavel (decisão do usuário de 19/10/2026: entra no limite do mês do crédito,
    que é a data do evento)."""
    chave = _chave_evento(livro, e)
    pos = _origem_do_evento(livro, e, chave)
    if pos.quantidade < 0:
        raise MotorError(f"{_nome(e)} em posição vendida")
    quantidade = e.quantidade or ZERO
    baixa = pos.vender(quantidade, e.valor)
    acao = e.tipo_ativo in (TipoAtivo.ACAO, TipoAtivo.UNIT)
    isentavel = acao and params.booleano("leilao_fracao_isentavel", e.data)
    livro.registrar(chave, e.data, f"evento:{e.id}",
                    "leilão de fração: venda pelo valor líquido recebido - PM x quantidade",
                    de=(e.conta, e.corretora), categoria=_categoria(e.tipo_ativo, False),
                    cv="V", quantidade=quantidade, valor=e.valor, baixa=baixa,
                    venda_acao=e.valor if isentavel else ZERO, isentavel=isentavel)  # fmt: skip


CUSTODIA = (
    "custódia inicial em 31/12: entra pelo custo de aquisição (IN RFB 1.585/2015 art. 58; "
    "o da declaração), sem resultado"
)


def _evento(livro: Livro, e: Evento) -> None:
    de = (e.conta, e.corretora)
    chave = _chave_evento(livro, e)
    pos = livro.posicoes[chave]
    if e.tipo is TipoEvento.TRANSFERENCIA:
        if e.conta_destino is None or e.corretora_destino is None:
            raise MotorError(f"transferência {e.ticker} ({e.data}) sem conta de destino")
        para = (e.conta_destino, e.corretora_destino)
        destino = livro.chave(e.titular, para, e.ativo, chave[3])
        if destino == chave:
            # mesma posição nesta visão (por titular; por corretora, entre contas da mesma
            # corretora): nada muda no PM nem no resultado, só a conta onde a posição está
            if e.quantidade is None:
                livro.origens[chave].discard(de)
            livro.origens[chave].add(para)
            return
        quantidade = pos.quantidade if e.quantidade is None else e.quantidade
        if quantidade <= 0 or quantidade > pos.quantidade:
            raise MotorError(
                f"transferência de {quantidade} {e.ticker} em {e.data}: posição na conta de "
                f"origem é {pos.quantidade}"
            )
        isento = (min(pos.isento_quantidade, quantidade * pos.isento_quantidade / pos.quantidade)
                  if pos.isento_quantidade else ZERO)  # fmt: skip
        pme_ate, media = pos.pme_ate, pos.pme_media
        custo_isento = pos.tirar_isento(isento) if isento else ZERO
        custo = pos.baixar_sem_resultado(quantidade)
        livro.registrar(
            chave,
            e.data,
            f"evento:{e.id}",
            "transferência de custódia (saída)",
            de=de,
            categoria=Categoria.COMUM,
            cv="V",
            quantidade=quantidade,
            valor=custo,
        )
        livro.posicoes[destino].comprar(quantidade, custo)
        if isento and pme_ate is not None:
            livro.posicoes[destino].somar_isento(isento, custo_isento, pme_ate, media)
        livro.registrar(
            destino,
            e.data,
            f"evento:{e.id}",
            "transferência de custódia (entrada, com o mesmo PM)",
            de=para,
            categoria=Categoria.COMUM,
            cv="C",
            quantidade=quantidade,
            valor=custo,
        )
        return
    if e.cv is None:
        raise MotorError(f"evento {e.tipo} de {e.ticker} ({e.data}) sem C/V")
    sinal = 1 if e.cv == "C" else -1
    quantidade = (e.quantidade or ZERO) * sinal
    custo = e.valor * sinal
    if e.tipo is TipoEvento.CUSTODIA_INICIAL and (
        e.cv != "C" or e.tipo_ativo in (TipoAtivo.OPCAO, TipoAtivo.FUTURO)
    ):
        # decisão do usuário (25/10/2026): só posição comprada à vista; opção, futuro e posição
        # vendida em aberto pedem as notas desde a abertura (prêmio, ajustes diários)
        raise MotorError(f"custódia inicial de {e.ticker} ({e.data}): só compra de ações, ETF, "
                         "BDR, units e FII; para opções, futuros e posição vendida, importe as "
                         "notas desde a abertura")  # fmt: skip
    abre = (e.tipo in ABREM_POSICAO or e.natureza is NaturezaEvento.OUTRA_CLASSE) and e.cv == "C"
    if pos.quantidade == 0 and not abre:
        # só a subscrição, a oferta pública, a custódia inicial e a bonificação noutra classe abrem
        # posição; nos demais o evento cairia numa posição que não existe
        raise MotorError(
            f"evento {e.tipo.value} de {e.ticker} ({e.data}) sem posição na conta "
            f"{e.corretora} {e.conta} (confira data, conta e ticker do evento)"
        )
    if pos.quantidade < 0 and (e.tipo in (TipoEvento.OFERTA_PUBLICA, TipoEvento.SUBSCRICAO)
                               or e.natureza is NaturezaEvento.OUTRA_CLASSE):  # fmt: skip
        # o ajuste abaixo zeraria a posição vendida e o ganho da venda sumiria sem resultado (regra
        # 5). A venda logo antes da liquidação da oferta já foi tratada em _antecipar_ofertas
        raise MotorError(
            f"{e.tipo.value.lower()} de {e.ticker} ({e.data}): há venda antes desta data e a "
            f"posição está vendida em {pos.quantidade} (confira a data do evento e as vendas)"
        )
    if e.tipo is TipoEvento.REDUCAO_CAPITAL and pos.custo + custo < 0 < pos.quantidade:
        raise MotorError(f"redução de capital de {e.ticker} ({e.data}) maior que o custo")
    if pos.quantidade > 0 > pos.quantidade + quantidade or pos.quantidade < 0 < (
        pos.quantidade + quantidade
    ):
        raise MotorError(
            f"evento {e.tipo.value} de {e.ticker} ({e.data}) inverte a posição de "
            f"{pos.quantidade} para {pos.quantidade + quantidade}"
        )
    if e.tipo in (TipoEvento.DESDOBRAMENTO, TipoEvento.GRUPAMENTO, TipoEvento.REDUCAO_CAPITAL):
        pos.repartir_isento(quantidade, custo)  # não é aquisição (IN 1.585 art. 58 §7º II)
    pos.ajustar(quantidade, custo)
    if (e.cv == "C" and e.pme_ate is not None and e.data <= e.pme_ate
            and e.tipo in (TipoEvento.BONIFICACAO, TipoEvento.SUBSCRICAO,
                           TipoEvento.CUSTODIA_INICIAL, TipoEvento.OFERTA_PUBLICA)):  # fmt: skip
        # IN 1.585 art. 67 IV e art. 69-A §1º: a data do evento decide o lote
        pos.somar_isento(quantidade, custo, e.pme_ate, e.pme_media_fim)
    if e.aviso and livro.visao is Visao.TITULAR and e.aviso not in livro.avisos:
        livro.avisos.append(e.aviso)
    regra = (
        CUSTODIA
        if e.tipo is TipoEvento.CUSTODIA_INICIAL
        else f"evento {e.tipo.value.lower()}: ajusta quantidade e custo, sem resultado"
    )
    livro.registrar(chave, e.data, f"evento:{e.id}", regra, de=de, categoria=Categoria.COMUM,
                    cv=e.cv, quantidade=quantidade, valor=custo, evento=e.tipo)  # fmt: skip
