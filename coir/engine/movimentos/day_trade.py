# pyright: reportPrivateUsage=false, reportUnusedFunction=false
# (os nomes com _ são internos do pacote, usados entre os arquivos dele)
"""Day trade: compra e venda do mesmo ativo no mesmo dia e corretora, casadas pela ordem dos
negócios (IN RFB 1.585/2015 art. 65 §3º); o excedente vira operação comum."""

from collections import defaultdict
from collections.abc import Sequence
from datetime import date
from decimal import Decimal

from coir.domain.fiscal import (
    ZERO,
    TipoAtivo,
)
from coir.engine.movimentos.livro import (
    MERCADOS_ACAO_A_VISTA,
    MERCADOS_DT,
    MERCADOS_EXERCICIO,
    Livro,
    OrigemMovimento,
    _categoria,
    _de,
    _Op,
    _Peca,
    _proporcao,
)
from coir.engine.posicao import Baixa
from coir.erros import MotorError


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
    # afasta o limite de R$ 20 mil; critério do COIR de 15/10/2026)
    pme = all(p.base.op.pme_ate is not None and dia <= p.base.op.pme_ate for p in compras)
    # a venda de ação em day trade entra na soma do limite de R$ 20 mil (parâmetro
    # day_trade_na_soma_limite, ligado: leitura literal de "total das alienações" da Lei 11.033
    # art. 3º I e da IN 1.585 art. 59 I; o §2º I só tira o day trade da isenção, e só a PME tem
    # regra de não computar, art. 67 §3º), mas o ganho dele nunca é isento
    acao = tipo in (TipoAtivo.ACAO, TipoAtivo.UNIT) and primeiro.op.mercado in MERCADOS_ACAO_A_VISTA
    na_soma = acao and not pme and livro.parametros.booleano("day_trade_na_soma_limite", dia)
    livro.registrar(chave, dia, origem,
                    f"{regra}: resultado = vendas casadas - compras casadas", de=de,
                    toca_posicao=False,
                    categoria=_categoria(tipo, day_trade=True), cv=None, quantidade=dt,
                    valor=alienacao, custos=custos, isento_pme=pme,
                    venda_acao=sum((p.bruto for p in vendas), ZERO) if na_soma else ZERO,
                    baixa=Baixa(resultado=alienacao - custo, custo_baixa=custo))  # fmt: skip
