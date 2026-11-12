"""Os fatos do domínio como eventos do registro de auditoria (``coir.auditoria``): os serviços
chamam uma linha, e o que vai para o registro se decide aqui. Quantidade, preço e valor exatos não
vão."""

import hashlib
from collections.abc import Sequence
from datetime import date
from decimal import Decimal
from typing import TYPE_CHECKING

from coir import auditoria
from coir.domain.fiscal import Apuracao, Categoria, MesApurado
from coir.domain.nota import NotaExtraida
from coir.erros import ConciliacaoError
from coir.parsing.extrair import Rejeitada

if TYPE_CHECKING:
    from coir.servicos.conciliacao import Conciliacao
    from coir.servicos.darf import LinhaDarf
    from coir.servicos.eventos import DadosEvento, Proposta


def nota(n: NotaExtraida, situacao: str, operacoes: int = 0) -> None:
    auditoria.evento(
        "importacao.nota", arquivo=n.arquivo, nota=n.numero, conta=n.conta, corretora=n.corretora,
        segmento=n.segmento.value, data=n.data_pregao, negocios=len(n.negocios_bovespa) + len(
            n.negocios_bmf), operacoes=operacoes, situacao=situacao,
    )  # fmt: skip


def rejeicao(r: Rejeitada) -> None:
    """Por que a nota não entrou: classe e ponto do erro e, na conciliação, campo e diferença."""
    if not auditoria.ativa():
        return
    conciliacao_ = r.erro if isinstance(r.erro, ConciliacaoError) else None
    diferenca = None
    if (
        conciliacao_
        and isinstance(conciliacao_.lido, Decimal)
        and isinstance(conciliacao_.esperado, Decimal)
    ):
        diferenca = conciliacao_.lido - conciliacao_.esperado
    auditoria.evento(
        "importacao.rejeicao", "A", arquivo=r.arquivo, paginas=r.paginas, nota=r.nota or None,
        data=r.data, campo=conciliacao_.campo if conciliacao_ else None, diferenca=diferenca,
        **(auditoria.dados_do_erro(r.erro) if r.erro else {}),
    )  # fmt: skip


def _mes(m: MesApurado) -> None:
    assinatura = hashlib.sha256(m.model_dump_json().encode()).hexdigest()
    if not auditoria.uma_vez("mes", m.titular, m.ano, m.mes, assinatura):
        return  # o mês não mudou desde o último cálculo nesta sessão
    if not (m.imposto or m.darf or m.darf_acumulado or m.irrf_mes or m.provisorio) and not any(
        c.resultado or c.vendas_limite or c.prejuizo_a_compensar for c in m.categorias
    ):
        return  # mês sem nada (o motor apura todos desde a primeira nota)
    por = {c.categoria: c for c in m.categorias}
    comum, day_trade, fii = (por.get(c) for c in (Categoria.COMUM, Categoria.DAY_TRADE,
                                                  Categoria.FII))  # fmt: skip
    auditoria.evento(
        "apuracao.mes", titular=m.titular, mes=date(m.ano, m.mes, 1), assinatura=assinatura,
        resultado_comum=comum.resultado if comum else None,
        resultado_day_trade=day_trade.resultado if day_trade else None,
        resultado_fii=fii.resultado if fii else None,
        vendas_acoes=comum.vendas_acoes if comum else None,
        vendas_limite=comum.vendas_limite if comum else None,
        isento=bool(comum and comum.isento > 0),
        prejuizo=sum((c.prejuizo_a_compensar for c in m.categorias), Decimal(0)),
        imposto=m.imposto, irrf=m.irrf_mes, darf=m.darf, darf_acumulado=m.darf_acumulado,
        provisorio=m.provisorio, incompletos=m.incompletos, dependentes=m.dependentes,
    )  # fmt: skip


def apuracao(ap: Apuracao) -> None:
    """Grandeza e sinal de cada mês, mais uma assinatura: mostra o que mudou entre dois cálculos
    (depois de uma versão nova, por exemplo) sem mostrar o valor."""
    if not auditoria.ativa():
        return
    auditoria.evento("apuracao.resumo", meses=len(ap.meses), avisos=len(ap.avisos),
                     isolados=len(ap.isolados))  # fmt: skip
    for m in ap.meses:
        _mes(m)
    for i in ap.isolados:
        if auditoria.uma_vez("isolado", i.titular, i.ticker, i.desde):
            auditoria.evento("apuracao.isolado", "A", titular=i.titular, ativo=i.ticker,
                             categoria=i.categoria.value, desde=i.desde)  # fmt: skip


def conciliacao(titular: int, res: "Conciliacao") -> None:
    auditoria.evento("conciliacao.resumo", titular=titular, data=res.data, linhas=len(res.linhas),
                     divergentes=len(res.divergentes), avisos=len(res.avisos))  # fmt: skip
    for x in res.divergentes:  # a diferença pode ser a posição inteira: só a grandeza
        auditoria.evento("conciliacao.divergencia", "A", titular=titular, data=res.data,
                         corretora=x.corretora, ativo=x.codigo, quantidade=x.diferenca)  # fmt: skip


def propostas(titular: int, lista: Sequence["Proposta"]) -> None:
    """Quantas por situação e, uma vez por sessão, cada proposta (tipo, ativo e data)."""
    if not auditoria.ativa():
        return
    situacoes = [p.situacao for p in lista]
    auditoria.evento(
        "propostas.resumo", titular=titular, total=len(lista),
        pendentes=situacoes.count("pendente"), ignoradas=situacoes.count("ignorada"),
        lancadas=situacoes.count("lançado"),
    )  # fmt: skip
    for p in lista:
        if auditoria.uma_vez("proposta", titular, p.chave, p.situacao):
            auditoria.evento(
                "eventos.proposta", titular=titular, tipo=p.dados.tipo.value,
                ativo=p.dados.ticker or p.dados.ticker_destino or None, data=p.data,
                origem=p.dados.origem, situacao=p.situacao, corretora=p.corretora,
            )  # fmt: skip


def evento_salvo(titular: int, d: "DadosEvento", alterado: bool) -> None:
    auditoria.evento(
        "eventos.salvo", titular=titular, tipo=d.tipo.value, ativo=d.ticker or None,
        ativo_destino=d.ticker_destino or None, data=d.data, origem=d.origem, alterado=alterado,
    )  # fmt: skip


def darfs(titular: int, linhas: Sequence["LinhaDarf"]) -> None:
    if not auditoria.ativa():
        return
    for x in linhas:
        if auditoria.uma_vez("darf", titular, x.ano, x.mes, x.situacao, x.principal):
            auditoria.evento(
                "darf.linha", titular=titular, mes=date(x.ano, x.mes, 1), situacao=x.situacao.value,
                valor=x.principal, com_multa=x.multa is not None, com_juros=x.juros is not None,
                com_pendencia=bool(x.pendencia),
            )  # fmt: skip
