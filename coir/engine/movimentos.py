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

from collections import defaultdict
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

from coir.domain.fiscal import (
    ZERO,
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
        self.movimentos: list[Movimento] = []
        self.avisos: list[str] = []

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
) -> list[Movimento]:
    """Movimentos de posição até a data `ate` (inclusive) na visão pedida (a apuração usa só a
    por titular). Situações aceitas mas que pedem conferência (venda a descoberto) vão para
    `avisos`."""
    livro = _executar(notas, eventos, parametros, ate, visao)
    if avisos is not None:
        avisos += livro.avisos
    return livro.movimentos


def carteira(
    notas: Iterable[NotaMotor],
    eventos: Iterable[Evento],
    parametros: ParametrosFiscais,
    ate: date,
    visao: Visao = Visao.TITULAR,
) -> list[PosicaoAtivo]:
    """Posições abertas no fim do dia `ate`. Por titular é o custo médio fiscal (apuração e
    IRPF); por corretora e por conta é a visão operacional, que não entra no imposto. Compras a
    termo ainda não liquidadas não estão na carteira."""
    livro = _executar(notas, eventos, parametros, ate, visao)
    saida: list[PosicaoAtivo] = []
    for (titular, grupo, ativo, vencimento), pos in livro.posicoes.items():
        if not pos.quantidade:
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
) -> Livro:
    livro = Livro(visao)
    dias: dict[date, _Dia] = defaultdict(_Dia)
    vencimentos: set[date] = set()
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
                    dias[op.liquidacao_termo].termos.append(item)
                continue
            dias[nota.data].operacoes.append(item)
            if op.tipo is TipoAtivo.OPCAO:
                if op.vencimento is None or op.vencimento < nota.data:
                    raise MotorError(
                        f"opção {op.ticker} em {nota.data}: vencimento {op.vencimento} ausente ou "
                        "anterior ao negócio (confira data/vencimentos_opcoes*.csv)"
                    )
                vencimentos.add(op.vencimento)
    for evento in eventos:
        if evento.data <= ate:
            livro.tickers.setdefault(evento.ativo, evento.ticker)
            livro.tipos.setdefault(evento.ativo, evento.tipo_ativo)
            if evento.ativo_destino is not None and evento.ticker_destino is not None:
                livro.tickers.setdefault(evento.ativo_destino, evento.ticker_destino)
                livro.tipos.setdefault(evento.ativo_destino,
                                       evento.tipo_ativo_destino or evento.tipo_ativo)  # fmt: skip
            dias[evento.data].eventos.append(evento)
    for venc in vencimentos:
        if venc <= ate:
            dias.setdefault(venc, _Dia())  # o pó acontece mesmo sem negócio no dia
    for dia in sorted(dias):
        try:
            _processar_dia(livro, dia, dias[dia], parametros)
        except MotorError as erro:
            erro.data = erro.data or dia
            raise
    return livro


def _processar_dia(livro: Livro, dia: date, d: _Dia, params: ParametrosFiscais) -> None:
    # cisões depois dos demais eventos do dia (agrupadas por posição de origem e destino)
    cisoes: dict[tuple[Chave, int | None], list[Evento]] = defaultdict(list)
    for evento in sorted(d.eventos, key=lambda e: e.id):
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
        if origens.count(chave) > 1:  # o percentual é sobre o custo de antes de todas
            raise MotorError(f"cisão de {grupo[0].ticker} ({grupo[0].data}) para dois ativos no "
                             "mesmo dia (não implementado)")  # fmt: skip
        _cisao(livro, chave, grupo)
    livro.termo_hoje = {}
    for t in sorted(d.termos, key=lambda t: (t.nota.id, t.op.ordem)):
        chave = livro.chave(t.nota.titular, _de(t), t.op.ativo)
        pos = livro.posicoes[chave]
        _lote_pme(pos, t.op, dia)
        # ações vendidas antes da liquidação: a posição ficou vendida e o termo a cobre (resultado
        # na liquidação, limite de isenção do mês da venda, como na recompra da venda a descoberto)
        mes_limite = pos.descoberto_desde if pos.quantidade < 0 else None
        antes = pos.quantidade
        baixa = pos.comprar(t.op.quantidade, t.valor)
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
    compras, vendas, exercicios = _day_trade(livro, dia, d.operacoes, mesma_corretora,
                                             pela_marcacao)  # fmt: skip
    for p in sorted(compras, key=_ordem):
        _aplicar(livro, dia, p)
    for o in sorted(exercicios, key=lambda o: (o.nota.id, o.op.ordem)):
        _exercicio(livro, dia, o)
    for p in sorted(vendas, key=_ordem):
        _aplicar(livro, dia, p)
    for chave, pos in list(livro.posicoes.items()):
        if pos.quantidade and chave[3] == dia:
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
                sum((o.op.quantidade for o in mc), ZERO)
                == dt
                == sum((o.op.quantidade for o in mv), ZERO)
            ):
                # os negócios marcados D primeiro: são exatamente o day trade
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


def _aplicar_parte(
    livro: Livro, dia: date, chave: Chave, p: _Peca, custo_termo: Decimal | None = None
) -> None:
    """Aplica a peça à posição. Com `custo_termo`, é a venda de ações do lote do termo na data da
    liquidação: sai pelo preço a termo, vai para a linha "Mercado a termo" da ficha do IRPF e nunca
    é isenta (IN 1.585 art. 62 I e art. 59 §2º IV; decisão do usuário de 20/10/2026)."""
    o = p.base
    pos = livro.posicoes[chave]
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
        _lote_pme(pos, o.op, dia)
        baixa = pos.comprar(p.quantidade, p.valor)
        venda_acao = ZERO
        if pos.quantidade >= 0:
            pos.descoberto_desde = None
    else:
        ja_vendida = pos.quantidade < 0
        # Lei 13.043/2014 art. 16 (IN 1.585 arts. 66 a 69-A): ganho isento, fora do limite
        pme = pos.quantidade > 0 and pos.pme_ate is not None and dia <= pos.pme_ate
        if pme and p.quantidade > pos.quantidade:
            raise MotorError(f"venda de {o.op.ticker} ({dia}) além do lote isento de PME")
        if custo_termo is None:
            baixa = pos.vender(p.quantidade, p.valor)
        elif pos.quantidade >= p.quantidade:
            baixa = pos.vender_lote(p.quantidade, p.valor, custo_termo)
        else:
            raise MotorError(f"{dia}: venda de {o.op.ticker} na data da liquidação do termo maior "
                             f"que a posição ({pos.quantidade})")  # fmt: skip
        venda_acao = p.bruto if acao and not pme else ZERO
        if pos.quantidade == 0:
            pos.pme_ate = None
        if pos.quantidade < 0 and not ja_vendida:
            pos.descoberto_desde = dia
    # venda a descoberto é operação normal: sem aviso (decisão do usuário de 25/10/2026)
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


def _lote_pme(pos: Posicao, op: OperacaoMotor, dia: date) -> None:
    """IN 1.585 art. 69-A §4º: lotes isentos (antes da oferta subsequente) e tributados têm
    controle separado. Sem os dois juntos no livro, a posição inteira é de um tipo só."""
    if pos.pme_ate is not None and dia > pos.pme_ate:
        pos.pme_ate = None  # isenção encerrada: o lote antigo já é tributado como o novo
    if pos.quantidade <= 0:
        pos.pme_ate = op.pme_ate
    elif pos.pme_ate != op.pme_ate:
        raise MotorError(
            f"{op.ticker} em {dia}: compra mistura lote isento de PME e lote tributado "
            "(controle separado do art. 69-A §4º não implementado)"
        )


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
    if pos.pme_ate is not None:
        raise MotorError(f"{_nome(e)}: lote isento de PME (controle separado não implementado)")
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
    abre = e.tipo in (TipoEvento.SUBSCRICAO, TipoEvento.CUSTODIA_INICIAL) and e.cv == "C"
    if pos.quantidade == 0 and not abre:
        # só a subscrição e a custódia inicial abrem posição; nos demais o evento cairia numa
        # posição que não existe
        raise MotorError(
            f"evento {e.tipo.value} de {e.ticker} ({e.data}) sem posição na conta "
            f"{e.corretora} {e.conta} (confira data, conta e ticker do evento)"
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
    pos.ajustar(quantidade, custo)
    regra = (
        CUSTODIA
        if e.tipo is TipoEvento.CUSTODIA_INICIAL
        else f"evento {e.tipo.value.lower()}: ajusta quantidade e custo, sem resultado"
    )
    livro.registrar(chave, e.data, f"evento:{e.id}", regra, de=de, categoria=Categoria.COMUM,
                    cv=e.cv, quantidade=quantidade, valor=custo, evento=e.tipo)  # fmt: skip
