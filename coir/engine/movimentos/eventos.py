# pyright: reportPrivateUsage=false, reportUnusedFunction=false
# (os nomes com _ são internos do pacote, usados entre os arquivos dele)
"""Eventos corporativos e de custódia na posição: desdobramento, grupamento, bonificação,
incorporação, cisão, redução de capital, direitos, leilão de fração e transferências."""

from datetime import date
from decimal import Decimal

from coir.domain.fiscal import (
    ZERO,
    Categoria,
    Evento,
    NaturezaEvento,
    TipoAtivo,
    TipoEvento,
    Visao,
)
from coir.domain.formato import reais
from coir.domain.parametros import ParametrosFiscais
from coir.engine.movimentos.livro import Chave, Livro, _categoria
from coir.engine.posicao import Posicao
from coir.erros import MotorError

ABREM_POSICAO = frozenset(
    {TipoEvento.SUBSCRICAO, TipoEvento.OFERTA_PUBLICA, TipoEvento.CUSTODIA_INICIAL,
     TipoEvento.DIREITO_RECEBIDO}
)  # fmt: skip


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
    if e.natureza is NaturezaEvento.RECIBO_SUBSCRICAO:
        # o recibo vira a ação ou a cota: não há alienação pelo titular (Lei 7.713/1988 art. 3º
        # §3º); o custo integral segue com ela (IN 1.585 art. 58 §§4º e 6º, por analogia)
        regra = "recibo de subscrição convertido (IN 1.585 art. 58 §§4º e 6º, por analogia): "
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


def _dividendo_em_bens(livro: Livro, e: Evento, params: ParametrosFiscais) -> None:
    """Dividendo pago em bens (ações ou cotas de outra empresa): o bem entra na posição na data da
    entrega. O dividendo é rendimento isento pelo valor atribuído no aviso (Lei 9.249/1995 art.
    10), fora da apuração mensal. Custo do bem: não houve valor pago, e a norma manda usar o valor
    corrente na data da aquisição (IN RFB 1.585/2015 art. 58 §3º IV; Lei 7.713/1988 art. 16 V);
    o COIR fica com o menor entre o atribuído e o corrente (parâmetro custo_dividendo_em_bens,
    conservador: custo menor, ganho maior na venda), com aviso quando o corrente é o menor
    (parecer do especialista e decisão de 16/11/2026). A posição na companhia que pagou não
    muda."""
    if not e.quantidade_destino or e.quantidade_destino <= 0 or e.valor <= 0:
        raise MotorError(f"{_nome(e)} sem a quantidade recebida ou o valor atribuído")
    if e.valor_corrente is None or e.valor_corrente <= 0:
        raise MotorError(f"{_nome(e)} sem o valor de mercado dos bens na entrega")
    criterio = params.texto("custo_dividendo_em_bens", e.data)
    if criterio not in ("menor_valor", "valor_atribuido"):
        raise MotorError(f"parâmetro custo_dividendo_em_bens: valor {criterio!r} desconhecido")
    de = (e.conta, e.corretora)
    destino = _destino(livro, e, livro.chave(e.titular, de, e.ativo), 1)
    custo = e.valor
    regra = "custo pelo valor atribuído no aviso"
    if criterio == "menor_valor" and e.valor_corrente < e.valor:
        custo = e.valor_corrente
        regra = "custo pelo valor de mercado na entrega, menor que o atribuído no aviso"
        aviso = (
            f"{e.data:%d/%m/%Y}: dividendo de {e.ticker} pago em {e.ticker_destino}: o custo "
            f"ficou pelo valor de mercado na entrega ({reais(custo)}), menor que o atribuído no "
            f"aviso ({reais(e.valor)}), que continua sendo o rendimento isento. É a leitura "
            "conservadora (IN RFB 1.585/2015 art. 58 §3º IV); o parâmetro "
            "custo_dividendo_em_bens muda o critério"
        )
        if livro.visao is Visao.TITULAR and aviso not in livro.avisos:
            livro.avisos.append(aviso)
    livro.posicoes[destino].comprar(e.quantidade_destino, custo)
    livro.registrar(destino, e.data, f"evento:{e.id}",
                    f"dividendo de {e.ticker} pago em bens (Lei 9.249/1995 art. 10): {regra}, sem "
                    "resultado", de=de, categoria=Categoria.COMUM, cv="C",
                    quantidade=e.quantidade_destino, valor=custo, evento=e.tipo)  # fmt: skip


def _amortizacao_fip(livro: Livro, e: Evento, pos: Posicao, custo: Decimal) -> Decimal:
    """Amortização de cotas de FIP comum: quanto do valor lançado reduz o custo (negativo).

    Até 31/12/2023 (parâmetro fip_amortizacao_custo = excedente): o imposto na fonte só incide
    sobre o que exceder o custo (IN RFB 1.585/2015 art. 32 §2º; Lei 11.312/2006 art. 2º §2º): a
    amortização reduz o custo até zero, e o excedente é rendimento tributado na fonte.
    Desde 01/01/2024 (proporcional_vp): só a parcela custo x amortização / valor patrimonial é
    devolução de custo; o restante é rendimento com IRRF de 15%, definitivo (Lei 14.754/2023 art.
    24 caput e §2º c/c art. 17 §2º III e §5º II b; art. 32 I). O COIR não tem o valor patrimonial:
    o usuário lança a parcela de custo do informe; com o valor inteiro, o custo cai até zero
    (custo menor, imposto maior na venda: nunca a menos), sempre com aviso (critério do COIR,
    conservador; parecer do especialista de 13/11/2026)."""
    regime = {p.valor for p in livro.parametros.parametros
              if p.nome == "fip_amortizacao_custo" and p.vigente(e.data)}  # fmt: skip
    avisar = livro.visao is Visao.TITULAR
    if pos.custo + custo < 0 < pos.quantidade:
        excedente = -(pos.custo + custo)
        custo = -pos.custo
        if avisar:
            livro.avisos.append(
                f"{e.data:%d/%m/%Y}: amortização de {e.ticker} acima do custo: o custo foi a zero "
                f"e o excedente de {reais(excedente)} é rendimento do fundo, tributado na fonte "
                "pelo administrador (IN RFB 1.585/2015 art. 32 §2º; Lei 14.754/2023 art. 24): "
                "declare pelo informe, na ficha de tributação exclusiva, pelo valor líquido"
            )
    if regime == {"proporcional_vp"} and avisar:
        livro.avisos.append(
            f"{e.data:%d/%m/%Y}: amortização de {e.ticker}: desde 2024 só parte do valor "
            "amortizado devolve custo (o custo vezes a amortização dividida pelo valor "
            "patrimonial da cota; Lei 14.754/2023 art. 24 §2º); o restante é rendimento com IRRF "
            "de 15%, definitivo. O COIR reduziu o custo pelo valor lançado no evento: lance só a "
            "parcela de custo que o informe do administrador mostra. Com o valor inteiro, o "
            "custo fica menor e o imposto na venda, maior"
        )
    return custo


def _regra(livro: Livro, nome: str, dia: date) -> str | None:
    """Valor do parâmetro de critério na data (None sem vigência: o motor isola, sem palpite)."""
    return livro.parametros.regra(nome, dia)


def _amortizacao_etf(livro: Livro, e: Evento, pos: Posicao, custo: Decimal) -> Decimal:
    """Amortização de cotas de ETF de ações desde 01/01/2024 (parâmetro etf_amortizacao_custo =
    proporcional_vp): só a parcela custo x amortização / valor patrimonial devolve custo; o
    restante é rendimento com IRRF de 15%, definitivo (Lei 14.754/2023 arts. 18 II, 24 caput e §2º
    c/c art. 17 §2º III e §5º II b; art. 32 I). O COIR não tem o valor patrimonial: o usuário lança
    a parcela de custo do informe; com o valor inteiro, o custo cai até zero (imposto na venda
    maior, nunca menor), com aviso. Como o FIP desde 2024 (parecer do especialista de
    24/11/2026)."""
    avisar = livro.visao is Visao.TITULAR
    if pos.custo + custo < 0 < pos.quantidade:
        excedente = -(pos.custo + custo)
        custo = -pos.custo
        if avisar:
            livro.avisos.append(
                f"{e.data:%d/%m/%Y}: amortização de {e.ticker} acima do custo: o custo foi a zero "
                f"e o excedente de {reais(excedente)} é rendimento do ETF, com IRRF de 15% "
                "definitivo (Lei 14.754/2023 arts. 18 II e 24): declare pelo informe, na ficha de "
                "tributação exclusiva (06), pelo valor líquido"
            )  # fmt: skip
    if avisar:
        livro.avisos.append(
            f"{e.data:%d/%m/%Y}: amortização de {e.ticker}: desde 2024 só a parcela de custo (o "
            "custo vezes a amortização dividida pelo valor patrimonial da cota; Lei 14.754/2023 "
            "arts. 24 §2º e 17 §5º II b) reduz o custo; o restante é rendimento com IRRF de 15%, "
            "definitivo. O COIR reduziu o custo pelo valor lançado no evento: lance só a parcela "
            "de custo que o informe do administrador mostra. Com o valor inteiro, o custo fica "
            "menor e o imposto na venda, maior"
        )  # fmt: skip
    return custo


def _acima_do_custo(livro: Livro, e: Evento, pos: Posicao, custo: Decimal) -> Decimal:
    """Restituição de capital ou amortização maior que o custo da posição: custo e preço médio a
    zero (nunca custo negativo: ele levaria o excedente para a apuração da bolsa, com a isenção e
    a compensação), e o excedente só é avisado (parecer do especialista de 24/11/2026, decisões do
    usuário). O custo fiscal é o do titular (IN RFB 1.585/2015 art. 58 caput): só a visão TITULAR
    decide e avisa; nas outras, o custo vai a zero sem aviso.

    - FI-Infra e FIP-IE: rendimento do fundo, alíquota zero (Lei 11.478/2007 art. 2º §2º; Lei
      12.431/2011 art. 3º; critério do COIR de 10/11/2026).
    - Ação e unit (restituicao_capital_acima_do_custo = ganho_capital): ganho de capital fora de
      bolsa, no GCAP, no mês do crédito do dinheiro (Lei 7.713/1988 art. 3º; IN 1.585 art. 58 §8º;
      SC Cosit 37/2020 por analogia); com isolar, o ativo sai do cálculo.
    - FII e Fiagro (fii_amortizacao = reduz_custo): excedente tributado a 20% na fonte pelo
      administrador, por analogia com o resgate (Lei 8.668/1993 art. 18 I; IN 1.585 art. 37 §3º).
    - ETF de renda fixa: tributação só na fonte (Lei 13.043/2014 art. 2º §4º).
    - Os demais (BDR, ETF antes de 2024, direito...): sem norma, o ativo fica isolado."""
    excedente = -(pos.custo + custo)
    zerado = -pos.custo
    if livro.visao is not Visao.TITULAR:
        return zerado
    tipo, dia = e.tipo_ativo, f"{e.data:%d/%m/%Y}"
    regra_acao = _regra(livro, "restituicao_capital_acima_do_custo", e.data)
    if tipo is TipoAtivo.FUNDO_INFRA:
        aviso = (
            f"{dia}: amortização de {e.ticker} acima do custo: o custo foi a zero e o excedente "
            f"de {reais(excedente)} é rendimento do fundo (Lei 11.478/2007 art. 2º §2º; Lei "
            "12.431/2011 art. 3º), com alíquota zero para a pessoa física: declare na mesma linha "
            "dos rendimentos do fundo (exclusiva 06)"
        )  # fmt: skip
    elif tipo in (TipoAtivo.ACAO, TipoAtivo.UNIT) and regra_acao == "ganho_capital":
        aviso = (
            f"{dia}: restituição de capital de {e.ticker} acima do custo: o custo e o preço médio "
            f"foram a zero, e o excedente de {reais(excedente)} ({reais(-custo)} recebidos - "
            f"{reais(pos.custo)} de custo) não entra na apuração mensal da bolsa. Na leitura do "
            "COIR, é ganho de capital fora de bolsa (Lei 7.713/1988 art. 3º; IN RFB 1.585/2015 "
            "art. 58 §8º; SC Cosit 37/2020, por analogia). Apure-o no programa GCAP no mês em que "
            "o dinheiro foi creditado; o DARF (código 4600) vence no último dia útil do mês "
            "seguinte. O ganho não compensa prejuízo da bolsa (SC Cosit 130/2023). Até R$ 35 mil "
            "recebidos no mês, pode ser isento (Lei 9.250/1995 art. 22 II). Antes, confira se o "
            "custo no COIR está completo (custódia inicial, bonificações com custo atribuído). "
            "Não há norma expressa: confirme com o contador"
        )  # fmt: skip
    elif tipo is TipoAtivo.FII and _regra(livro, "fii_amortizacao", e.data) == "reduz_custo":
        aviso = (
            f"{dia}: amortização de {e.ticker} acima do custo: o custo foi a zero, e o excedente "
            f"de {reais(excedente)} não entra na apuração mensal. Não há norma expressa para a "
            "amortização de FII. Por analogia com o resgate (Lei 8.668/1993 art. 18 I; IN RFB "
            "1.585/2015 art. 37 §3º), o excedente é tributado a 20% na fonte pelo administrador. "
            "Confira o informe; se não houve retenção, fale com o administrador ou com o contador"
        )  # fmt: skip
    elif tipo is TipoAtivo.ETF_RF:
        aviso = (
            f"{dia}: amortização de {e.ticker} acima do custo: o custo foi a zero. O ETF de renda "
            "fixa é tributado só na fonte (Lei 13.043/2014 art. 2º §4º): o excedente de "
            f"{reais(excedente)} é rendimento retido pelo administrador. Declare pelo informe, na "
            "ficha de tributação exclusiva (06), pelo líquido"
        )  # fmt: skip
    else:
        raise MotorError(
            f"restituição de capital de {e.ticker} ({dia}) maior que o custo: não há norma para a "
            "pessoa física neste tipo de ativo, e o COIR não calcula. Confira o informe do "
            "emissor ou do depositário e fale com o contador. Para seguir, corrija o custo "
            "(custódia inicial) ou o valor do evento"
        )  # fmt: skip
    livro.avisos.append(aviso)
    return zerado


def _leilao_fracao(livro: Livro, e: Evento, params: ParametrosFiscais) -> None:
    """Fração vendida em leilão pela companhia: alienação pelo valor líquido recebido, custo pelo
    PM (IN RFB 1.585/2015 art. 58 caput). Sem norma sobre a isenção: parâmetro
    leilao_fracao_isentavel (critério do COIR de 19/10/2026: entra no limite do mês do crédito,
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


def _direito_nao_exercido(livro: Livro, e: Evento) -> None:
    """Direito, recibo ou bônus de subscrição comprado que venceu sem exercício nem venda: o custo
    pago (com corretagem e emolumentos, pelo PM) vira perda em operações comuns na data do evento,
    o fim do prazo de exercício (IN RFB 1.585/2015 art. 60 §3º, por analogia com a opção que vira
    pó; art. 64 caput; SC Cosit 166/2021; parecer do especialista de 09/11/2026). Fora do limite
    de R$ 20 mil: não há valor de alienação. Baixa o menor entre a quantidade do evento e a
    posição: o que passa da posição são direitos recebidos sem custo, que nunca entraram no livro
    e não dão resultado; o que sobra da posição continua em aberto, com aviso. Sem quantidade (a
    linha "não exercido" da B3 vem sem ela), baixa todo o saldo."""
    chave = _chave_evento(livro, e)
    pos = livro.posicoes[chave]
    if pos.quantidade < 0:
        raise MotorError(f"{_nome(e)} em posição vendida")
    quantidade = pos.quantidade if e.quantidade is None else min(e.quantidade, pos.quantidade)
    if quantidade <= 0:
        return  # só direitos recebidos sem custo: nada a baixar
    baixa = pos.vender(quantidade, ZERO)
    livro.registrar(chave, e.data, f"evento:{e.id}",
                    "direito de subscrição não exercido no prazo: o custo pago vira perda (IN RFB "
                    "1.585/2015 art. 60 §3º, por analogia)", de=(e.conta, e.corretora),
                    categoria=Categoria.COMUM, cv=None, quantidade=quantidade, valor=ZERO,
                    baixa=baixa, evento=e.tipo)  # fmt: skip
    if pos.quantidade > 0 and livro.visao is Visao.TITULAR:
        livro.avisos.append(
            f"{e.data}: {e.ticker}: o evento de direito não exercido baixou {quantidade}, e "
            f"{pos.quantidade} continuam na posição (confira a quantidade na movimentação da B3)"
        )


DIREITO_RECEBIDO = (
    "direito de subscrição recebido sem custo, na data do crédito (IN RFB 1.585/2015 art. 58 §7º, "
    "por analogia; Lei 7.713/1988 art. 16 §4º): entra no custo médio dos direitos com custo zero"
)


def _exercicio_direito(livro: Livro, e: Evento) -> None:
    """Direito de subscrição exercido: os direitos saem pelo custo médio, sem resultado, e o
    recibo (ou a ação) subscrito entra pelo valor pago mais o custo dos direitos usados (Lei
    7.713/1988 art. 16 caput e §2º; IN RFB 1.585/2015 art. 58 caput; art. 60 II a e §1º, por
    analogia com a opção de compra exercida). O exercício não é alienação: sem resultado, fora do
    limite de R$ 20 mil e sem IRRF. Os direitos que faltarem na posição foram recebidos sem custo
    e não estavam no livro: entram com custo zero, com aviso."""
    if not e.quantidade or e.quantidade <= 0:
        raise MotorError(f"{_nome(e)} sem a quantidade de direitos exercidos")
    if not e.quantidade_destino or e.quantidade_destino <= 0:
        raise MotorError(f"{_nome(e)} sem a quantidade subscrita")
    if e.valor <= 0:
        raise MotorError(f"{_nome(e)} sem o valor pago na subscrição")
    chave = _chave_evento(livro, e)
    pos = livro.posicoes[chave]
    if pos.quantidade < 0:
        raise MotorError(f"{_nome(e)} em posição vendida")
    de = (e.conta, e.corretora)
    falta = e.quantidade - pos.quantidade
    if falta > 0:
        pos.comprar(falta, ZERO)
        livro.registrar(chave, e.data, f"evento:{e.id}",
                        "direito recebido sem custo: entra antes do exercício, com custo zero (IN "
                        "RFB 1.585/2015 art. 58 §7º, por analogia)", de=de,
                        categoria=Categoria.COMUM, cv="C", quantidade=falta,
                        valor=ZERO)  # fmt: skip
        if livro.visao is Visao.TITULAR:
            livro.avisos.append(
                f"{e.data:%d/%m/%Y}: {e.ticker}: {falta} direito(s) exercido(s) não estavam na "
                "posição e entraram como recebidos sem custo; se foram comprados, falta a nota "
                "(confira a movimentação da B3)"
            )
    destino = _destino(livro, e, chave, 1)
    custo = pos.baixar_sem_resultado(e.quantidade)
    regra = "direito de subscrição exercido (IN 1.585 art. 58 caput e art. 60 II a, por analogia): "
    livro.registrar(chave, e.data, f"evento:{e.id}",
                    regra + f"sai para {e.ticker_destino} com o custo, sem resultado", de=de,
                    categoria=Categoria.COMUM, cv="V", quantidade=e.quantidade, valor=custo,
                    evento=e.tipo)  # fmt: skip
    livro.posicoes[destino].ajustar(e.quantidade_destino, custo + e.valor)
    livro.registrar(destino, e.data, f"evento:{e.id}",
                    regra + f"entra pelo valor pago ({reais(e.valor)}) mais o custo dos direitos "
                    f"({reais(custo)})", de=de, categoria=Categoria.COMUM, cv="C",
                    quantidade=e.quantidade_destino, valor=custo + e.valor,
                    evento=e.tipo)  # fmt: skip


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
        # critério do COIR (25/10/2026): só posição comprada à vista; opção, futuro e posição
        # vendida em aberto pedem as notas desde a abertura (prêmio, ajustes diários)
        raise MotorError(f"custódia inicial de {e.ticker} ({e.data}): só compra de ações, ETF, "
                         "BDR, units e FII; para opções, futuros e posição vendida, importe as "
                         "notas desde a abertura")  # fmt: skip
    if e.tipo is TipoEvento.DIREITO_RECEBIDO and (e.cv != "C" or e.valor):
        raise MotorError(f"{_nome(e)}: o direito recebido entra (C) com custo zero; o comprado "
                         "vem na nota de corretagem")  # fmt: skip
    if e.tipo is TipoEvento.DIREITO_RECEBIDO and any(
        m.ativo == e.ativo and m.titular == e.titular and m.origem.startswith("direito:")
        for m in livro.movimentos
    ):
        # uma venda anterior já pôs esses direitos no livro, a custo zero: somar de novo deixaria
        # direitos que não existem na carteira
        raise MotorError(f"{_nome(e)}: há venda deste direito antes da data do crédito, e os "
                         "direitos recebidos já entraram nessa venda. Use no evento a data em que "
                         "a B3 creditou os direitos, anterior à primeira venda")  # fmt: skip
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
    if e.tipo is TipoEvento.REDUCAO_CAPITAL and e.tipo_ativo is TipoAtivo.FIP:
        custo = _amortizacao_fip(livro, e, pos, custo)
    elif (e.tipo is TipoEvento.REDUCAO_CAPITAL and e.tipo_ativo is TipoAtivo.ETF
          and _regra(livro, "etf_amortizacao_custo", e.data) == "proporcional_vp"):  # fmt: skip
        custo = _amortizacao_etf(livro, e, pos, custo)
    elif e.tipo is TipoEvento.REDUCAO_CAPITAL and pos.custo + custo < 0 < pos.quantidade:
        custo = _acima_do_custo(livro, e, pos, custo)
    if pos.quantidade > 0 > pos.quantidade + quantidade or pos.quantidade < 0 < (
        pos.quantidade + quantidade
    ):
        raise MotorError(
            f"evento {e.tipo.value} de {e.ticker} ({e.data}) inverte a posição de "
            f"{pos.quantidade} para {pos.quantidade + quantidade}"
        )
    if e.tipo in (TipoEvento.DESDOBRAMENTO, TipoEvento.GRUPAMENTO, TipoEvento.REDUCAO_CAPITAL):
        # não é aquisição: desdobramento e grupamento (IN 1.585 art. 58 §7º II); redução de
        # capital (art. 58 §8º)
        pos.repartir_isento(quantidade, custo)
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
        else DIREITO_RECEBIDO
        if e.tipo is TipoEvento.DIREITO_RECEBIDO
        else f"evento {e.tipo.value.lower()}: ajusta quantidade e custo, sem resultado"
    )
    livro.registrar(chave, e.data, f"evento:{e.id}", regra, de=de, categoria=Categoria.COMUM,
                    cv=e.cv, quantidade=quantidade, valor=custo, evento=e.tipo)  # fmt: skip
