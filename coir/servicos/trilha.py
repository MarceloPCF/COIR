"""Os fatos do domínio como eventos do registro de auditoria (``coir.auditoria``).

Fica aqui para os serviços continuarem legíveis: cada um chama uma linha e o que vai para o
registro (e o que não vai: quantidade, preço e valor exatos) se decide neste arquivo. O motor não
registra nada; a apuração é registrada a partir do resultado dele.
"""

import hashlib
import logging
import re
from collections.abc import Iterable, Sequence
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import TYPE_CHECKING

from sqlalchemy import func, select
from sqlalchemy.exc import SQLAlchemyError

from coir import auditoria, referencias
from coir.db.modelos import (
    ArquivoImportado,
    Conta,
    EventoCorporativo,
    NotaCorretagem,
    Operacao,
    Titular,
)
from coir.db.sessao import conectar, revisao_final
from coir.domain.fiscal import Apuracao, Categoria, MesApurado
from coir.erros import CoirError, ConciliacaoError
from coir.parsing.extrair import Rejeitada, ResultadoExtracao

log = logging.getLogger(__name__)

if TYPE_CHECKING:
    from coir.servicos.darf import LinhaDarf
    from coir.servicos.eventos import Proposta

RE_MUTAVEL = re.compile(r"«[^»]*»|[\d.,/:\-]+")  # a página muda; a forma do layout, não


def arquivo_lendo(arquivo: str, tamanho: int) -> None:
    """Antes de ler: se o programa travar ou cair, o último arquivo fica registrado."""
    auditoria.evento("importacao.arquivo", arquivo=arquivo, tamanho=tamanho, situacao="lendo")


def arquivo_ignorado(arquivo: str, tamanho: int) -> None:
    auditoria.evento("importacao.arquivo", arquivo=arquivo, tamanho=tamanho, situacao="repetido")


def extracao(arquivo: str, res: ResultadoExtracao) -> None:
    """O que o leitor viu em cada página e cada nota lida de um arquivo."""
    if not auditoria.ativa():
        return
    for p in res.leituras:
        auditoria.evento(
            "extracao.pagina", arquivo=arquivo, pagina=p.pagina, corretora=p.corretora,
            segmento=p.segmento, layout=p.layout, tipo=p.tipo, folha=p.folha, nota=p.nota,
            data=p.data,
        )  # fmt: skip
    for n in res.notas:
        negocios = len(n.negocios_bovespa) + len(n.negocios_bmf)
        auditoria.evento(
            "extracao.nota", arquivo=arquivo, nota=n.numero, conta=n.conta, corretora=n.corretora,
            segmento=n.segmento.value, data=n.data_pregao, paginas=n.paginas, negocios=negocios,
        )  # fmt: skip
    auditoria.evento(
        "importacao.arquivo", arquivo=arquivo, situacao="lido", total=len(res.leituras),
        notas=len(res.notas), rejeitadas=len(res.rejeitadas), avisos=len(res.avisos),
    )  # fmt: skip


def rejeicao(r: Rejeitada) -> None:
    """Por que a nota não entrou: a classe do erro e, na conciliação, o campo e a diferença."""
    if not auditoria.ativa():
        return
    campo: str | None = None
    diferenca: Decimal | None = None
    if isinstance(r.erro, ConciliacaoError):
        campo = r.erro.campo
        if isinstance(r.erro.esperado, Decimal) and isinstance(r.erro.lido, Decimal):
            diferenca = r.erro.lido - r.erro.esperado
    auditoria.evento(
        "extracao.rejeicao", "A", arquivo=r.arquivo, paginas=r.paginas, nota=r.nota or None,
        data=r.data, classe=type(r.erro).__name__ if r.erro else None, campo=campo,
        diferenca=diferenca, texto=r.motivo, lido=getattr(r.erro, "impresso", None),
    )  # fmt: skip
    if r.esqueleto and auditoria.uma_vez("esqueleto", RE_MUTAVEL.sub("#", "".join(r.esqueleto))):
        auditoria.evento(
            "extracao.esqueleto", arquivo=r.arquivo, pagina=r.paginas[0] if r.paginas else None,
            esqueleto=r.esqueleto,
        )  # fmt: skip


def avisos(textos: Iterable[str], evento: str = "importacao.aviso", **campos: object) -> None:
    for texto in textos:
        auditoria.evento(evento, "A", texto=texto, **campos)


def nota(arquivo: str, numero: str, situacao: str, operacoes: int = 0) -> None:
    auditoria.evento("importacao.nota", arquivo=arquivo, nota=numero, situacao=situacao,
                     operacoes=operacoes)  # fmt: skip


def planilha(arquivo: str, tipo: object, linhas: int, novas: int, detalhe: str) -> None:
    auditoria.evento(
        "b3.planilha", arquivo=arquivo, tipo=str(tipo) if tipo else None,
        situacao="lida" if tipo else "recusada", linhas=linhas, novas=novas,
        texto=detalhe or None,
    )  # fmt: skip


def _mes(m: MesApurado) -> None:
    assinatura = hashlib.sha256(m.model_dump_json().encode()).hexdigest()
    if not auditoria.uma_vez("mes", m.titular, m.ano, m.mes, assinatura):
        return  # o mês não mudou desde o último cálculo nesta sessão
    if not (m.imposto or m.darf or m.darf_acumulado or m.irrf_mes or m.provisorio) and not any(
        c.resultado or c.vendas_acoes or c.prejuizo_a_compensar for c in m.categorias
    ):
        return  # mês sem nada (o motor apura todos desde a primeira nota): não diz nada
    por_categoria = {c.categoria: c for c in m.categorias}

    def resultado(categoria: Categoria) -> Decimal | None:
        c = por_categoria.get(categoria)
        return c.resultado if c is not None else None

    comum = por_categoria.get(Categoria.COMUM)
    auditoria.evento(
        "apuracao.mes", titular=m.titular, mes=date(m.ano, m.mes, 1), assinatura=assinatura,
        resultado_comum=resultado(Categoria.COMUM),
        resultado_day_trade=resultado(Categoria.DAY_TRADE),
        resultado_fii=resultado(Categoria.FII),
        vendas_acoes=comum.vendas_acoes if comum is not None else None,
        isento=comum is not None and comum.isento > 0,
        prejuizo=sum((c.prejuizo_a_compensar for c in m.categorias), Decimal(0)),
        imposto=m.imposto, irrf=m.irrf_mes, darf=m.darf, darf_acumulado=m.darf_acumulado,
        provisorio=m.provisorio, incompletos=m.incompletos, dependentes=m.dependentes,
    )  # fmt: skip


def apuracao(ap: Apuracao, op: auditoria.Operacao) -> None:
    """Grandeza e sinal de cada mês, mais uma assinatura: mostra o que mudou entre dois cálculos
    (depois de uma versão nova, por exemplo) sem mostrar o valor."""
    op.dados(meses=len(ap.meses), avisos=len(ap.avisos), isolados=len(ap.isolados))
    if not auditoria.ativa():
        return
    for m in ap.meses:
        _mes(m)
    for texto in ap.avisos:
        if auditoria.uma_vez("aviso da apuração", texto):
            auditoria.evento("apuracao.aviso", "A", texto=texto)
    for i in ap.isolados:
        if auditoria.uma_vez("isolado", i.titular, i.ticker, i.desde, i.motivo):
            auditoria.evento(
                "apuracao.isolado", "A", titular=i.titular, ativo=i.ticker,
                categoria=i.categoria.value, desde=i.desde,
                meses_fora=[a * 100 + m for a, m in i.meses], motivo=i.motivo,
            )  # fmt: skip


def propostas(titular: int, lista: Sequence["Proposta"], op: auditoria.Operacao) -> None:
    """Quantas por situação e, uma vez por sessão, cada proposta (tipo, ativo e data)."""
    situacoes = [p.situacao for p in lista]
    op.dados(
        total=len(lista), pendentes=situacoes.count("pendente"),
        ignoradas=situacoes.count("ignorada"), lancadas=situacoes.count("lançado"),
    )  # fmt: skip
    if not auditoria.ativa():
        return
    for p in lista:
        if auditoria.uma_vez("proposta", titular, p.chave, p.situacao):
            auditoria.evento(
                "eventos.proposta", titular=titular, tipo=p.dados.tipo.value,
                ativo=p.dados.ticker or p.dados.ticker_destino or None, data=p.data,
                origem=p.dados.origem, situacao=p.situacao, corretora=p.corretora,
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


def resumo_do_banco(banco: Path, pasta_tabelas: Path) -> None:
    """O tamanho do que está no banco e a versão das tabelas, para o suporte situar o resto.
    Consultar o banco aqui não pode impedir o programa de abrir."""
    if not auditoria.ativa():
        return
    info = referencias.info_pacote(pasta_tabelas)
    try:
        with conectar(banco) as fabrica, fabrica() as s:
            quantos = {
                nome: s.scalar(select(func.count()).select_from(tabela)) or 0
                for nome, tabela in (
                    ("titulares", Titular), ("contas", Conta), ("notas", NotaCorretagem),
                    ("operacoes", Operacao), ("eventos", EventoCorporativo),
                    ("arquivos", ArquivoImportado),
                )
            }  # fmt: skip
    except (CoirError, SQLAlchemyError) as erro:
        log.warning("resumo do banco indisponível: %s", type(erro).__name__)
        return
    auditoria.evento("sessao.banco", revisao=str(revisao_final()),
                     tabelas=info[1] if info else None, **quantos)  # fmt: skip
