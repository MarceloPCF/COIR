# pyright: reportPrivateUsage=false, reportUnusedFunction=false
# (os nomes com _ são internos do pacote, usados entre os arquivos dele)
"""Entrada e saída de quantidade e custo numa posição: compra, venda pelo custo médio ou por lote
(PME) e exercício de opção."""

from datetime import date
from decimal import ROUND_FLOOR, Decimal

from coir.domain.fiscal import (
    ZERO,
    Categoria,
    Grupo,
    OperacaoMotor,
    TipoAtivo,
    Visao,
)
from coir.engine.movimentos.livro import (
    MERCADOS_ACAO_A_VISTA,
    MERCADOS_EXERCICIO,
    SEM_IRRF_DA_VENDA,
    Chave,
    Livro,
    _categoria,
    _de,
    _Op,
    _Peca,
    _proporcao,
)
from coir.engine.posicao import Posicao
from coir.erros import MotorError


def _aplicar(livro: Livro, dia: date, p: _Peca) -> None:
    o = p.base
    chave = livro.chave(o.nota.titular, _de(o), o.op.ativo, o.op.vencimento)
    lote = livro.termo_hoje.get(chave)
    if o.op.cv != "V" or o.op.mercado in MERCADOS_EXERCICIO or lote is None or not lote.quantidade:
        _aplicar_parte(livro, dia, chave, p)
        return
    # venda na data da liquidação do termo (critério do COIR de 21/10/2026, parecer do
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
    """IN 1.585 art. 69-A §4º: cada lote tem o seu PM. A IN não fixa qual sai primeiro
    (parâmetro pme_lote_ordem_venda). Padrão ``proporcional``: a venda se reparte entre os lotes na
    proporção das quantidades, como na SC Cosit 148/2025 (analogia; critério do COIR de
    10/11/2026, leitura conservadora), com a parte isenta arredondada para baixo na unidade da
    quantidade vendida. ``isento_primeiro`` e ``tributado_primeiro`` continuam como opções
    (decisão de 08/11/2026). O que passar da posição é venda a descoberto, como em qualquer ativo;
    a venda inteira se reparte pela quantidade."""
    pos = livro.posicoes[chave]
    antes = pos.quantidade
    isento, tributado = pos.isento_quantidade, pos.quantidade - pos.isento_quantidade
    ordem = livro.parametros.texto("pme_lote_ordem_venda", dia)
    if ordem == "proporcional":
        vendida = min(p.quantidade, antes)
        inteira = p.quantidade == p.quantidade.to_integral_value()
        unidade = Decimal(1) if inteira else Decimal("0.00000001")
        parte = (vendida * isento / antes).quantize(unidade, rounding=ROUND_FLOOR)
        lotes = [("isento", parte), ("tributado", min(vendida - parte, tributado))]
    elif ordem in ("isento_primeiro", "tributado_primeiro"):
        lotes = [("isento", isento), ("tributado", tributado)]
        if ordem == "tributado_primeiro":
            lotes.reverse()
    else:
        raise MotorError(f"parâmetro pme_lote_ordem_venda: valor {ordem!r} desconhecido")
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
    é isenta (IN 1.585 art. 62 I e art. 59 §2º IV; critério do COIR de 20/10/2026)."""
    o = p.base
    pos = livro.posicoes[chave]
    if (o.op.cv == "V" and custo_termo is None and lote is None and pos.quantidade > 0
            and pos.isento_em(dia)):  # fmt: skip
        _vender_por_lote(livro, dia, chave, p)
        return
    # ação ou unit no mercado à vista (units só entram com o parâmetro units_na_isencao)
    acao = o.op.tipo in (TipoAtivo.ACAO, TipoAtivo.UNIT) and o.op.mercado in MERCADOS_ACAO_A_VISTA
    # IN 1.585 art. 59 §2º IV: a isenção não se aplica à alienação por exercício de opção; o valor
    # continua na soma do limite (critério do COIR de 15/10/2026, leitura conservadora)
    isentavel = acao and o.op.mercado not in MERCADOS_EXERCICIO
    # venda a descoberto (leitura A1, critério do COIR de 15/10/2026, sem norma específica): o
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
    # nem ETF de renda fixa (tributação exclusiva); usada só para repartir o IRRF entre as fichas.
    # Em opções a base é a soma dos prêmios do dia, se positiva (IN RFB 1.585/2015 art. 63 II e
    # §1º II): a compra entra negativa e a soma por dia é feita na repartição
    base_irrf = p.bruto if (o.op.cv == "V" and o.op.mercado not in MERCADOS_EXERCICIO and not pme
                            and o.op.tipo not in SEM_IRRF_DA_VENDA) else ZERO  # fmt: skip
    if o.op.cv == "C" and o.op.tipo is TipoAtivo.OPCAO and o.op.mercado not in MERCADOS_EXERCICIO:
        base_irrf = -p.bruto
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
