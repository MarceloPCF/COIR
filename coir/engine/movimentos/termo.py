# pyright: reportPrivateUsage=false, reportUnusedFunction=false
# (os nomes com _ são internos do pacote, usados entre os arquivos dele)
"""Termo: compras a termo em aberto, venda a termo na liquidação e o aviso de venda a termo
descoberta."""

from collections.abc import Iterable, Sequence
from datetime import date, timedelta

from coir.domain.fiscal import (
    ZERO,
    Categoria,
    Grupo,
    NotaMotor,
    OperacaoMotor,
    TermoAberto,
    TipoAtivo,
)
from coir.domain.formato import reais
from coir.engine.movimentos.livro import Chave, Livro, _de, _Op
from coir.engine.rateio import ratear
from coir.erros import MotorError


def termos_em_aberto(notas: Iterable[NotaMotor], ate: date) -> list[TermoAberto]:
    """Compras e vendas a termo feitas até `ate` e ainda não liquidadas no fim desse dia. Não
    mudam a carteira nem o PM (entram na liquidação, ver `carteira`): são só para mostrar. Na
    compra, o custo é o que entrará no PM: valor a termo + custos da nota rateados como em
    `_percorrer`; na venda, o valor da alienação: valor a termo - custos."""
    saida: list[TermoAberto] = []
    for nota in notas:
        if nota.data > ate or not nota.operacoes:
            continue
        custos = ratear(nota.custos, [o.total for o in nota.operacoes])
        for op, custo in zip(nota.operacoes, custos, strict=True):
            if (op.mercado != "TERMO" or op.liquidacao_termo is None
                    or op.liquidacao_termo <= ate):  # fmt: skip
                continue
            saida.append(
                TermoAberto(
                    titular=nota.titular, corretora=nota.corretora, conta=nota.conta,
                    nota=nota.id, ativo=op.ativo, ticker=op.ticker, tipo=op.tipo,
                    pregao=nota.data, liquidacao=op.liquidacao_termo, quantidade=op.quantidade,
                    custo=op.total + custo if op.cv == "C" else op.total - custo, cv=op.cv,
                )
            )  # fmt: skip
    return saida


def _avisar_termo_descoberto(
    livro: Livro, compras: Sequence[tuple[NotaMotor, OperacaoMotor]]
) -> None:
    """Compra a termo sem a liquidação na movimentação da B3 (o serviço manda o aviso na operação)
    e venda à vista do mesmo ativo, entre o pregão e a data usada, que deixou a posição vendida:
    sem ações em custódia, a entrega da venda exigiu a liquidação antecipada do termo ("por
    diferença", de 1 a 3 dias úteis depois do pedido), e o ganho e o custo são da data dela (IN
    RFB 1.585/2015 art. 59 §2º IV e art. 62 I). O COIR mantém a data usada (o vencimento) e troca
    o aviso da operação por este, mais forte (parecer do especialista de 17/11/2026, opção (a),
    escolhida pelo usuário em 18/11/2026: sem palpite da data)."""
    for nota, op in compras:
        assert op.liquidacao_termo is not None
        venda = next((m for m in livro.movimentos
                      if (m.titular, m.ativo) == (nota.titular, op.ativo) and m.cv == "V"
                      and m.grupo is not Grupo.TERMO and m.quantidade_apos < 0
                      and nota.data <= m.data < op.liquidacao_termo), None)  # fmt: skip
        if venda is None:
            continue
        dia = venda.data
        uteis = [dia + timedelta(days=n) for n in range(1, 8)
                 if (dia + timedelta(days=n)).weekday() < 5][:3]  # fmt: skip
        meses = sorted({(d.year, d.month) for d in uteis})
        provavel = " ou ".join(f"{m:02d}/{a}" for a, m in meses)
        resultado = sum((m.resultado for m in livro.movimentos
                         if m.origem == f"termo:{op.id}"), ZERO)  # fmt: skip
        em_jogo = (
            f"o resultado de {reais(resultado)} que o COIR pôs em "
            f"{op.liquidacao_termo:%m/%Y} (e o imposto dele) provavelmente é de {provavel}"
            if resultado
            else f"o custo, o ganho e o imposto provavelmente são de {provavel}"
        )
        aviso = (
            f"Termo: a venda de {op.ticker} em {dia:%d/%m/%Y} deixou a posição vendida "
            f"({venda.quantidade_apos.normalize():f}) antes da data usada para a compra a termo "
            f"de {nota.data:%d/%m/%Y} ({op.liquidacao_termo:%d/%m/%Y}, sem a liquidação na "
            "movimentação da B3). Sem ações em custódia, a entrega da venda exigiu a liquidação "
            f"antecipada do termo, de 1 a 3 dias úteis depois da venda: {em_jogo}, com DARF "
            "que vence antes (IN RFB 1.585/2015 art. 59 §2º IV e art. 62 I). Importe a "
            "movimentação da B3 que cobre o período do contrato para o COIR ler a data da "
            "liquidação"
        )
        if op.aviso in livro.avisos:
            livro.avisos.remove(op.aviso)
        if aviso not in livro.avisos:
            livro.avisos.append(aviso)


def _venda_a_termo(livro: Livro, dia: date, chave: Chave, t: _Op) -> None:
    """Liquidação da venda a termo, no vencimento do contrato: é quando a alienação se dá (IN RFB
    1.585/2015 art. 59 §2º IV). Ganho = preço a termo x quantidade, menos os custos da nota da
    venda, menos o custo médio da posição na abertura deste dia (art. 62 II, vendedor descoberto,
    e III, vendedor coberto; art. 58): as compras feitas entre o contrato e a liquidação entram no
    custo médio. Linha "Mercado a termo" da ficha, 15%, nunca isento (art. 59 §2º IV); o valor
    bruto entra na soma do limite de R$ 20 mil do mês da liquidação, como no exercício e no day
    trade (leitura literal do art. 59 I). Sem ações para entregar, ou com lote isento de PME, o
    motor para: nada de palpite (parecer do especialista de 13/11/2026)."""
    op, pos = t.op, livro.posicoes[chave]
    onde = f"venda a termo de {op.ticker} (nota {t.nota.numero} de {t.nota.data:%d/%m/%Y})"
    if pos.quantidade < op.quantidade:
        raise MotorError(
            f"{onde}: faltam ações para entregar na liquidação do termo em {dia:%d/%m/%Y} "
            f"(contrato de {op.quantidade}, posição de {pos.quantidade}). Confira se falta uma "
            "nota de compra ou um evento, ou se o contrato foi liquidado antes do vencimento"
        )
    if pos.isento_em(dia):  # sem norma para a isenção de PME na entrega do termo e sem caso real
        raise MotorError(f"{onde}: lote isento de PME dentro do prazo da isenção")
    baixa = pos.vender(op.quantidade, t.valor)
    lote = livro.termo_hoje.get(chave)
    if lote is not None and lote.quantidade > pos.quantidade:
        # ações de uma compra a termo liquidada hoje foram entregues: saem do lote do dia
        lote.baixar_sem_resultado(lote.quantidade - pos.quantidade)
    acao = op.tipo in (TipoAtivo.ACAO, TipoAtivo.UNIT)
    livro.registrar(
        chave, dia, f"termo:{op.id}",
        "venda a termo liquidada: preço a termo - custos da venda - custo médio do dia da "
        "liquidação (IN 1.585 art. 62 II e III); nunca isento (art. 59 §2º IV)",
        de=_de(t), categoria=Categoria.COMUM, cv="V", quantidade=op.quantidade, valor=t.valor,
        custos=t.custos, baixa=baixa, venda_acao=op.total if acao else ZERO, isentavel=False,
        grupo=Grupo.TERMO,
    )  # fmt: skip
