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
- Termo: as ações entram na carteira na liquidação, pelo preço a termo + custos da nota.
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
        toca_posicao: bool = True,
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
                origem=origem, regra=regra, categoria=categoria, grupo=_grupo(tipo), tipo=tipo,
                cv=cv,  # type: ignore[arg-type]
                quantidade=quantidade, valor=valor, custos=custos, custo_baixa=b.custo_baixa,
                resultado=b.resultado, venda_acao_a_vista=venda_acao, isentavel=isentavel,
                mes_limite=mes_limite, isento_pme=isento_pme,
                quantidade_apos=pos.quantidade, custo_apos=pos.custo,
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
            dias[evento.data].eventos.append(evento)
    for venc in vencimentos:
        if venc <= ate:
            dias.setdefault(venc, _Dia())  # o pó acontece mesmo sem negócio no dia
    for dia in sorted(dias):
        _processar_dia(livro, dia, dias[dia], parametros)
    return livro


def _processar_dia(livro: Livro, dia: date, d: _Dia, params: ParametrosFiscais) -> None:
    for evento in sorted(d.eventos, key=lambda e: e.id):
        _evento(livro, evento)
    for t in sorted(d.termos, key=lambda t: (t.nota.id, t.op.ordem)):
        chave = livro.chave(t.nota.titular, _de(t), t.op.ativo)
        pos = livro.posicoes[chave]
        _lote_pme(pos, t.op, dia)
        # ações vendidas antes da liquidação: a posição ficou vendida e o termo a cobre (resultado
        # na liquidação, limite de isenção do mês da venda, como na recompra da venda a descoberto)
        mes_limite = pos.descoberto_desde if pos.quantidade < 0 else None
        baixa = pos.comprar(t.op.quantidade, t.valor)
        if pos.quantidade >= 0:
            pos.descoberto_desde = None
        regra = "termo liquidado: entra pelo preço a termo + custos da compra"
        if baixa.resultado or baixa.custo_baixa:
            regra += "; cobre a venda feita antes da liquidação (resultado nesta data)"
        acao = t.op.tipo in (TipoAtivo.ACAO, TipoAtivo.UNIT)
        livro.registrar(chave, dia, f"termo:{t.op.id}", regra, de=_de(t),
                        categoria=Categoria.COMUM, cv="C", quantidade=t.op.quantidade,
                        valor=t.valor, custos=t.custos, baixa=baixa, isentavel=acao,
                        mes_limite=mes_limite)  # fmt: skip
    mesma_corretora = params.booleano("day_trade_mesma_corretora", dia)
    compras, vendas, exercicios = _day_trade(livro, dia, d.operacoes, mesma_corretora)
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
    livro: Livro, dia: date, ops: Sequence[_Op], mesma_corretora: bool
) -> tuple[list[_Peca], list[_Peca], list[_Op]]:
    """Casa compras e vendas do mesmo ativo no dia pela ordem dos negócios; devolve as peças
    comuns (compras, vendas) e os exercícios (nunca day trade)."""
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
            raise MotorError(f"mercado {o.op.mercado} sem regra no motor")
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
        dt_c, sobra_c = _fatiar(c, dt, "compra")
        dt_v, sobra_v = _fatiar(v, dt, "venda")
        _movimento_day_trade(livro, dia, grupo, dt, dt_c, dt_v)
        compras += sobra_c
        vendas += sobra_v
    return compras, vendas, exercicios


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
    vendas: list[_Peca],
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
                    "day trade (IN 1.585 art. 65 §3º: 1ª compra com a 1ª venda, pela ordem): "
                    "resultado = vendas casadas - compras casadas", de=de, toca_posicao=False,
                    categoria=_categoria(tipo, day_trade=True), cv=None, quantidade=dt,
                    valor=alienacao, custos=custos, isento_pme=pme,
                    baixa=Baixa(resultado=alienacao - custo, custo_baixa=custo))  # fmt: skip


def _aplicar(livro: Livro, dia: date, p: _Peca) -> None:
    o = p.base
    chave = livro.chave(o.nota.titular, _de(o), o.op.ativo, o.op.vencimento)
    pos = livro.posicoes[chave]
    # ação ou unit no mercado à vista (units só entram com o parâmetro units_na_isencao)
    acao = o.op.tipo in (TipoAtivo.ACAO, TipoAtivo.UNIT) and o.op.mercado in MERCADOS_ACAO_A_VISTA
    # IN 1.585 art. 59 §2º IV: a isenção não se aplica à alienação por exercício de opção; o valor
    # continua na soma do limite (decisão do usuário de 15/10/2026, leitura conservadora)
    isentavel = acao and o.op.mercado not in MERCADOS_EXERCICIO
    descoberto = (
        o.op.cv == "V" and o.op.tipo is not TipoAtivo.OPCAO and pos.quantidade < p.quantidade
    )
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
        baixa = pos.vender(p.quantidade, p.valor)
        venda_acao = p.bruto if acao and not pme else ZERO
        if pos.quantidade == 0:
            pos.pme_ate = None
        if pos.quantidade < 0 and not ja_vendida:
            pos.descoberto_desde = dia
    aviso = f"{dia}: venda a descoberto de {o.op.ticker}"
    fiscal = livro.visao is Visao.TITULAR
    if descoberto and fiscal and not any(x.startswith(aviso) for x in livro.avisos):
        livro.avisos.append(
            f"{dia}: venda a descoberto de {o.op.ticker} (posição {pos.quantidade} após a venda): "
            "confira se falta transferência ou evento; o resultado sai na recompra"
        )
    livro.registrar(chave, dia, f"op:{o.op.id}", p.regra, de=_de(o),
                    categoria=_categoria(o.op.tipo, day_trade=False), cv=o.op.cv,
                    quantidade=p.quantidade, valor=p.valor, custos=p.custos, baixa=baixa,
                    venda_acao=venda_acao, isentavel=isentavel and not pme, mes_limite=mes_limite,
                    isento_pme=pme)  # fmt: skip


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


def _evento(livro: Livro, e: Evento) -> None:
    de = (e.conta, e.corretora)
    chave = livro.chave(e.titular, de, e.ativo)
    if e.tipo_ativo is TipoAtivo.OPCAO:  # evento na série (ajuste da B3): a posição aberta
        chave = livro.chave_opcao(e.titular, de, e.ativo, e.data) or chave
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
    if pos.quantidade == 0 and not (e.tipo is TipoEvento.SUBSCRICAO and e.cv == "C"):
        # só a subscrição abre posição; nos demais o evento cairia numa posição que não existe
        raise MotorError(
            f"evento {e.tipo.value} de {e.ticker} ({e.data}) sem posição na conta "
            f"{e.corretora} {e.conta} (confira data, conta e ticker no eventos.csv)"
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
    livro.registrar(chave, e.data, f"evento:{e.id}", f"evento {e.tipo.value.lower()}: ajusta "
                    "quantidade e custo, sem resultado", de=de, categoria=Categoria.COMUM, cv=e.cv,
                    quantidade=quantidade, valor=custo)  # fmt: skip
