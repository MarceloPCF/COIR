"""Banco + motor -> dados dos relatórios de um titular e ano -> arquivos (xlsx, csv, pdf).

A apuração vai até o fim do ano ou até a data de geração, o que vier antes: uma opção que
vence depois de hoje não pode virar pó num relatório do ano em curso.
"""

from datetime import date, datetime
from pathlib import Path

from sqlalchemy import select

from coir.config import dentro_do_repositorio
from coir.db import repositorio as repo
from coir.db.modelos import Conta
from coir.db.sessao import conectar
from coir.domain.fiscal import Visao
from coir.domain.parametros import ReferenciasFiscais
from coir.domain.referencias import Referencias
from coir.domain.relatorio import DadosRelatorio, NotaDoLivro
from coir.engine.apuracao import apurar
from coir.engine.movimentos import carteira
from coir.erros import ExportacaoError
from coir.reports.fiscais import TODOS, relatorios
from coir.reports.saida_csv import escrever_csv
from coir.reports.saida_pdf import escrever_pdf
from coir.reports.saida_xlsx import escrever_xlsx
from coir.servicos import outros_custos
from coir.servicos.apuracao import entradas
from coir.servicos.proventos import proventos_banco

FORMATOS = ("xlsx", "csv", "pdf")


def dados_banco(
    banco: Path, refs: Referencias, fiscais: ReferenciasFiscais, ano: int, gerado_em: datetime,
    titular: int | None = None, data: date | None = None, visao: Visao = Visao.TITULAR,
) -> DadosRelatorio:  # fmt: skip
    ate = min(date(ano, 12, 31), gerado_em.date())
    if ate.year != ano:
        raise ExportacaoError(f"ano {ano} ainda não começou")
    dia = data or ate
    with conectar(banco) as fabrica, fabrica() as sessao:
        todos = repo.titulares(sessao)
        if titular is None:
            if len(todos) != 1:
                ids = ", ".join(str(t.id) for t in todos) or "nenhum"
                raise ExportacaoError(f"escolha o titular com --titular (cadastrados: {ids})")
            escolhido = todos[0]
        else:
            escolhido = next((t for t in todos if t.id == titular), None)
            if escolhido is None:
                raise ExportacaoError(f"titular {titular} não cadastrado")
        notas, eventos = entradas(sessao, refs, fiscais)
        custos = outros_custos.do_banco(sessao)
        livro = tuple(NotaDoLivro(nota=n, operacoes=tuple(ops))
                      for n, ops in repo.carregar_notas(sessao, escolhido.id, ano))  # fmt: skip
        do_titular = select(Conta).where(Conta.titular_id == escolhido.id)
        contas = {c.id: c.numero for c in sessao.scalars(do_titular)}
        nome, cpf, tid = escolhido.nome, escolhido.cpf, escolhido.id
    lista_proventos, avisos_proventos = proventos_banco(banco, fiscais, tid)
    params = fiscais.parametros
    fiscal = carteira(notas, eventos, params, ate, Visao.TITULAR)
    anterior = carteira(notas, eventos, params, date(ano - 1, 12, 31), Visao.TITULAR)
    na_data = fiscal if (visao, dia) == (Visao.TITULAR, ate) else carteira(notas, eventos, params,
                                                                           dia, visao)  # fmt: skip
    return DadosRelatorio(
        titular=tid, nome=nome, cpf=cpf, ano=ano, gerado_em=gerado_em,
        codigo_darf=params.texto("darf_codigo", date(ano, 1, 1)),
        apuracao=apurar(notas, eventos, params, fiscais.calendario, ate, custos), ate=ate,
        data_carteira=dia,
        visao=visao, carteira=tuple(na_data), carteira_fiscal=tuple(fiscal),
        carteira_fiscal_anterior=tuple(anterior), contas=contas,
        notas=livro, proventos=tuple(lista_proventos), avisos_proventos=tuple(avisos_proventos),
    )  # fmt: skip


def validar(destino: Path, formatos: tuple[str, ...], chaves: tuple[str, ...] | None) -> None:
    """Destino e opções conferidos antes de carregar o banco e rodar o motor."""
    if dentro_do_repositorio(destino):
        raise ExportacaoError(f"destino dentro do repositório ({destino.resolve()}): os "
                              "relatórios têm dados pessoais; use a pasta de dados")  # fmt: skip
    errados = [f for f in formatos if f not in FORMATOS]
    if errados:
        raise ExportacaoError(f"formato desconhecido: {', '.join(errados)} (use {FORMATOS})")
    desconhecidos = [c for c in chaves or () if c not in TODOS]
    if desconhecidos:
        raise ExportacaoError(f"relatório desconhecido: {', '.join(desconhecidos)} "
                              f"(disponíveis: {', '.join(TODOS)})")  # fmt: skip


def exportar(
    dados: DadosRelatorio, destino: Path, formatos: tuple[str, ...] = FORMATOS,
    chaves: tuple[str, ...] | None = None, prefixo: str | None = None,
) -> list[Path]:  # fmt: skip
    """Grava os relatórios; o nome dos arquivos não leva CPF (só o cabeçalho, dentro deles)."""
    validar(destino, formatos, chaves)
    destino.mkdir(parents=True, exist_ok=True)
    prefixo = prefixo or f"coir-{dados.ano}-titular{dados.titular}"
    rels = relatorios(dados, chaves)
    saida: list[Path] = []
    if "xlsx" in formatos:
        saida.append(escrever_xlsx(rels, destino / f"{prefixo}.xlsx", dados.gerado_em))
    if "csv" in formatos:
        for r in rels:
            saida += escrever_csv(r, destino, prefixo)
    if "pdf" in formatos:
        saida += [escrever_pdf(r, destino / f"{prefixo}-{r.chave}.pdf") for r in rels if r.pdf]
    return saida


def exportar_carteira(
    banco: Path, refs: Referencias, fiscais: ReferenciasFiscais, titular: int, dia: date,
    visao: Visao, agora: datetime, destino: Path, formatos: tuple[str, ...] = FORMATOS,
) -> list[Path]:  # fmt: skip
    """A carteira de uma data e visão (tela Carteira), em arquivos próprios: não sobrescreve os
    relatórios do ano."""
    validar(destino, formatos, ("carteira",))
    dados = dados_banco(banco, refs, fiscais, dia.year, agora, titular, dia, visao)
    prefixo = f"coir-carteira-{dia:%Y-%m-%d}-por-{visao.value}-titular{titular}"
    return exportar(dados, destino, formatos, ("carteira",), prefixo)
