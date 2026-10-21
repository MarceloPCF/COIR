"""Linha de comando do COIR (Typer)."""

import logging
import sys
from pathlib import Path
from typing import TYPE_CHECKING

import typer

from coir import __version__
from coir.config import VAR_DADOS, dados_dir
from coir.erros import CoirError

if TYPE_CHECKING:
    from datetime import datetime

    from coir.domain.parametros import ReferenciasFiscais
    from coir.domain.referencias import Referencias
    from coir.gabarito.modelo import Gabarito
    from coir.gabarito.pseudonimos import MapaPseudonimos

app = typer.Typer(
    help="COIR v3 — Controle de Operações e Imposto de Renda (B3).",
    no_args_is_help=True,
)


@app.callback()
def principal(
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Log detalhado."),
) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )


@app.command()
def versao() -> None:
    """Mostra a versão do COIR."""
    print(f"COIR {__version__}")


@app.command()
def config() -> None:
    """Mostra a pasta de dados configurada e se ela existe."""
    pasta = dados_dir()
    situacao = "existe" if pasta.is_dir() else "NÃO existe"
    print(f"Pasta de dados ({VAR_DADOS}): {pasta} [{situacao}]")


gabarito_app = typer.Typer(
    help="Gabarito: planilha do usuário convertida para validação.", no_args_is_help=True
)
app.add_typer(gabarito_app, name="gabarito")

OpcaoBanco = typer.Option(None, help="Arquivo do banco (padrão: coir.db na pasta de dados).")
OpcaoCaminho = typer.Option(None, help="Padrão: dentro da pasta de dados (COIR_DADOS).")
OpcaoTitularUnico = typer.Option(
    False, help="Trata todos os CPFs da planilha como um único titular."
)


def _origem(origem: Path | None) -> Path:
    from coir.gabarito.leitor import localizar_planilha

    return origem or localizar_planilha(dados_dir(exigir=True) / "_gabarito")


def _destino(destino: Path | None) -> Path:
    return destino or dados_dir(exigir=True) / "gabarito"


def _ler_planilha(
    origem: Path, destino: Path, titular_unico: bool = False
) -> tuple["Gabarito", "MapaPseudonimos"]:
    from coir.gabarito.exportar import ARQUIVO_MAPA, PASTA_PRIVADA
    from coir.gabarito.extrair import ABAS, extrair
    from coir.gabarito.leitor import ler_abas
    from coir.gabarito.pseudonimos import MapaPseudonimos

    mapa = MapaPseudonimos.carregar(destino / PASTA_PRIVADA / ARQUIVO_MAPA)
    return extrair(ler_abas(origem, ABAS), mapa, titular_unico), mapa


@gabarito_app.command("piloto")
def gabarito_piloto(
    notas: str = typer.Option(
        "", help="PDFs separados por vírgula (padrão: 8 notas de referência)."
    ),
    origem: Path | None = OpcaoCaminho,
    destino: Path | None = OpcaoCaminho,
    titular_unico: bool = OpcaoTitularUnico,
) -> None:
    """Gera piloto.md: gabarito x resumo do PDF para poucas notas."""
    from coir.gabarito import piloto
    from coir.gabarito.exportar import ARQUIVO_MAPA, PASTA_PRIVADA, exigir_fora_do_repositorio

    pasta = _destino(destino)
    exigir_fora_do_repositorio(pasta)
    gabarito, mapa = _ler_planilha(_origem(origem), pasta, titular_unico)
    mapa.salvar(pasta / PASTA_PRIVADA / ARQUIVO_MAPA)
    nomes = tuple(n.strip() for n in notas.split(",") if n.strip()) or piloto.NOTAS_PADRAO
    texto = piloto.gerar(gabarito, dados_dir(exigir=True) / "_notas-corretagens", nomes)
    saida = pasta / "piloto" / "piloto.md"
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(texto, encoding="utf-8")
    print(f"Piloto gravado em {saida} ({len(nomes)} PDF(s)).")


@gabarito_app.command("conferir")
def gabarito_conferir(
    origem: Path | None = OpcaoCaminho,
    destino: Path | None = OpcaoCaminho,
    titular_unico: bool = OpcaoTitularUnico,
) -> None:
    """Confere automaticamente o gabarito contra todos os PDFs de notas (conferencia.md/.csv)."""
    import csv

    from coir.gabarito import conferencia
    from coir.gabarito.exportar import exigir_fora_do_repositorio

    pasta = _destino(destino)
    exigir_fora_do_repositorio(pasta)
    gabarito, _ = _ler_planilha(_origem(origem), pasta, titular_unico)
    resultados = conferencia.conferir_pasta(gabarito, dados_dir(exigir=True) / "_notas-corretagens")
    md, linhas = conferencia.relatorio(resultados)
    saida = pasta / "conferencia"
    saida.mkdir(parents=True, exist_ok=True)
    (saida / "conferencia.md").write_text(md, encoding="utf-8")
    with (saida / "conferencia.csv").open("w", encoding="utf-8", newline="") as arquivo:
        csv.writer(arquivo, lineterminator="\n").writerows(linhas)
    print("\n".join(md.splitlines()[2:6]))
    print(f"Relatório em {saida}")


@gabarito_app.command("gerar")
def gabarito_gerar(
    origem: Path | None = OpcaoCaminho,
    destino: Path | None = OpcaoCaminho,
    titular_unico: bool = OpcaoTitularUnico,
) -> None:
    """Converte a planilha-gabarito em CSV pseudonimizados + manifesto + resumo.md."""
    from coir.gabarito import resumo
    from coir.gabarito.exportar import exportar

    arquivo, pasta = _origem(origem), _destino(destino)
    gabarito, mapa = _ler_planilha(arquivo, pasta, titular_unico)
    manifesto = exportar(gabarito, mapa, arquivo, pasta)
    (pasta / "resumo.md").write_text(resumo.gerar(gabarito), encoding="utf-8")
    print(f"Gabarito gravado em {pasta}")
    print(f"Contagens: {manifesto['contagens']}")
    for aviso in gabarito.avisos:
        print(f"Aviso: {aviso}")


@gabarito_app.command("resumo")
def gabarito_resumo(pasta: Path | None = OpcaoCaminho) -> None:
    """Mostra o resumo a partir dos CSV já gerados."""
    from coir.gabarito import resumo
    from coir.gabarito.exportar import carregar

    print(resumo.gerar(carregar(_destino(pasta))))


@gabarito_app.command("externo")
def gabarito_externo(
    arquivo: Path | None = typer.Argument(
        None,
        help="xlsx de operações finalizadas de outro sistema (padrão: o mais recente "
        "'*Finalizadas*.xlsx' em _relatorios-b3).",
    ),
    visao: str = typer.Option("titular", help="Visão do motor: titular, corretora ou conta."),
    desde: str | None = typer.Option(
        None, help="Compara a partir de AAAA-MM (marco, ex.: mês seguinte à carteira zerada)."
    ),
    banco: Path | None = OpcaoBanco,
    destino: Path | None = OpcaoCaminho,
) -> None:
    """Compara as vendas fechadas de uma planilha externa com o motor (dia, corretora, ativo)."""
    import csv
    from collections import Counter
    from datetime import datetime
    from decimal import Decimal

    from coir.domain.fiscal import Visao
    from coir.gabarito.comparar import Tolerancias
    from coir.gabarito.exportar import exigir_fora_do_repositorio
    from coir.gabarito.externo import (
        carregar_explicacoes,
        comparar_externo,
        explicar,
        ler_planilha_externa,
        ticker_atual,
    )
    from coir.servicos.apuracao import movimentos_banco

    dados = dados_dir(exigir=True)
    if arquivo is None:
        candidatos = sorted((dados / "_relatorios-b3").glob("*Finalizadas*.xlsx"))
        if not candidatos:
            raise typer.BadParameter("nenhum '*Finalizadas*.xlsx' em _relatorios-b3")
        arquivo = candidatos[-1]
    try:
        vista = Visao(visao)
    except ValueError as erro:
        raise typer.BadParameter(str(erro)) from erro
    marco = None
    if desde is not None:
        try:
            marco = datetime.strptime(desde, "%Y-%m").date()
        except ValueError as erro:
            raise typer.BadParameter(f"--desde deve ser AAAA-MM (lido {desde!r})") from erro
    refs, tol = _refs(), Tolerancias.carregar()
    vendas = ler_planilha_externa(arquivo)
    movs = movimentos_banco(_banco(banco), refs, _fiscais(),
                            max(v.data for v in vendas), vista)  # fmt: skip
    linhas, meses = comparar_externo(vendas, movs, ticker_atual(refs), tol.operacao, tol.mes,
                                    marco)  # fmt: skip
    conhecidas = _destino(destino) / "divergencias-externo.csv"
    linhas = explicar(linhas, carregar_explicacoes(conhecidas))
    contagem = Counter(x.classe for x in linhas)
    grandes = sum(1 for x in linhas if x.classe == "diferença" and abs(x.diferenca) > 1)
    print(f"Planilha externa: {len(vendas)} vendas; {len(linhas)} chaves (dia, corretora, ativo); "
          f"visão do motor: {vista.value}"
          + (f"; desde {marco:%m/%Y}" if marco else ""))  # fmt: skip
    for classe, n in contagem.most_common():
        extra = f" ({grandes} acima de R$ 1,00)" if classe == "diferença" else ""
        print(f"  {classe}: {n}{extra}")
    print(f"Meses x mercado com diferença acima de R$ {tol.mes}: {len(meses)}")
    for m in meses:
        aberta = sum((x.diferenca for x in linhas if (x.data.year, x.data.month) == (m.ano, m.mes)
                      and x.classe not in ("confere", "explicada")), Decimal(0))  # fmt: skip
        nota = " - explicado venda a venda" if abs(aberta) <= tol.mes else ""
        print(f"  {m.mes:02d}/{m.ano} {m.mercado}: motor {m.motor:.2f} externo {m.externo:.2f} "
              f"(dif {m.motor - m.externo:.2f}){nota}")  # fmt: skip
    pasta = _destino(destino) / "conferencia"
    exigir_fora_do_repositorio(pasta)
    pasta.mkdir(parents=True, exist_ok=True)
    saida = pasta / f"externo-{vista.value}.csv"
    with saida.open("w", encoding="utf-8", newline="") as f:
        escrita = csv.writer(f, lineterminator="\n")
        escrita.writerow(
            ["data", "corretora", "ticker", "motor", "externo", "diferenca", "classe", "motivo"]
        )
        for x in linhas:
            if x.classe != "confere":
                escrita.writerow([x.data, x.corretora, x.ticker, f"{x.motor:.2f}",
                                  f"{x.externo:.2f}", f"{x.diferenca:.2f}", x.classe,
                                  x.motivo])  # fmt: skip
    print(f"Detalhe em {saida} (motivos já analisados: {conhecidas})")


@gabarito_app.command("dirpf")
def gabarito_dirpf(
    ano: int = typer.Option(..., help="Ano-calendário da declaração entregue (ex.: 2025)."),
    pdf: Path | None = typer.Option(
        None, help="PDF 'Imprimir > Declaração' do programa IRPF (padrão: o único em _dirpf)."
    ),
    titular: int | None = typer.Option(None, help="Id do titular (coir titular listar)."),
    banco: Path | None = OpcaoBanco,
    destino: Path | None = OpcaoCaminho,
) -> None:
    """Compara as fichas do COIR com a declaração entregue (critério de saída da F8)."""
    from collections import Counter

    from coir import referencias
    from coir.gabarito.comparar import Tolerancias
    from coir.gabarito.dirpf import ler_dirpf
    from coir.gabarito.dirpf_comparar import carregar_explicacoes, comparar, explicar, gravar
    from coir.gabarito.exportar import exigir_fora_do_repositorio
    from coir.gabarito.externo import ticker_atual
    from coir.reports.irpf import fichas
    from coir.servicos.declaracao import declaracoes_lidas
    from coir.servicos.relatorios import dados_banco

    dados = dados_dir(exigir=True)
    if pdf is None:
        lidas, avisos = declaracoes_lidas(dados)
        do_ano = [d for _, d in lidas if d.ano == ano]
        if len(do_ano) > 1:  # vários titulares: a do titular escolhido
            cpf = dados_banco(_banco(banco), _refs(), _fiscais(), ano, _agora(), titular).cpf
            do_ano = [d for d in do_ano if d.cpf == cpf]
        if len(do_ano) != 1:
            raise typer.BadParameter(f"{len(do_ano)} DIRPF do ano-calendário {ano} em _dirpf "
                                     f"({'; '.join(avisos) or 'informe --pdf'})")  # fmt: skip
        declarada = do_ano[0]
    else:
        declarada = ler_dirpf(pdf)
    if declarada.ano != ano:
        raise typer.BadParameter(f"o PDF é do ano-calendário {declarada.ano}, não de {ano}")
    pasta_refs = referencias.pasta_efetiva(dados)
    d = dados_banco(_banco(banco), _refs(), _fiscais(), ano, _agora(), titular,
                    pasta_refs=pasta_refs)  # fmt: skip
    geradas = fichas(d)
    if geradas is None or d.irpf is None:
        raise typer.BadParameter(f"sem mapeamento das fichas do IRPF {ano + 1} em data/irpf")
    tol = Tolerancias.carregar()
    conhecidas = _destino(destino) / "dirpf-explicadas.csv"
    linhas = explicar(comparar(geradas, declarada, d.irpf, tol.mes, tol.operacao,
                              ticker_atual(_refs())),
                      carregar_explicacoes(conhecidas))  # fmt: skip
    print(f"DIRPF {declarada.exercicio} (ano-calendário {ano}) x COIR, titular {d.titular}")
    for secao in ("renda_variavel", "fii", "bens", "rendimentos", "imposto_pago"):
        contagem = Counter(x.classe for x in linhas if x.secao == secao)
        print(f"  {secao}: " + ", ".join(f"{c} {n}" for c, n in contagem.most_common()))
    abertas = [x for x in linhas if x.classe not in ("confere", "explicada")]
    print(f"Divergências sem explicação: {len(abertas)}")
    if declarada.sem_fonte:
        print("  Linhas da DIRPF sem leitura por fonte (só o total): "
              + ", ".join(f"{f} {c}" for f, c in declarada.sem_fonte))  # fmt: skip
    pasta = _destino(destino)
    exigir_fora_do_repositorio(pasta)
    saida = pasta / "dirpf-divergencias.csv"
    gravar([x for x in linhas if x.classe != "confere"], saida)
    print(f"Detalhe em {saida} (motivos já analisados: {conhecidas})")


@app.command("validar-gabarito")
def validar_gabarito(
    candidato: Path = typer.Option(..., help="Pasta com CSV no formato do gabarito."),
    gabarito: Path | None = OpcaoCaminho,
    niveis: str = typer.Option(
        "operacoes,posicoes,apuracao",
        help="Níveis separados por vírgula (operacoes, negocios, notas, posicoes, apuracao).",
    ),
    explicacoes: Path | None = typer.Option(
        None, help="CSV de divergências conhecidas (padrão: na pasta do gabarito)."
    ),
    relatorio: Path | None = typer.Option(None, help="Grava todas as divergências neste arquivo."),
) -> None:
    """Compara um candidato com o gabarito usando as tolerâncias de data/tolerancias.toml."""
    from coir.gabarito.comparar import (
        Tolerancias,
        carregar_explicacoes,
        com_tickers_da_data,
        comparar,
        explicacao,
    )
    from coir.gabarito.exportar import carregar

    pasta = _destino(gabarito)
    conhecidas = carregar_explicacoes(explicacoes or pasta / "divergencias-conhecidas.csv")
    rel = comparar(
        com_tickers_da_data(carregar(pasta), _refs()),
        carregar(candidato),
        Tolerancias.carregar(),
        [n.strip() for n in niveis.split(",") if n.strip()],
    )
    linhas = [
        f"{d.nivel} {d.chave} {d.campo}: esperado {d.esperado} obtido {d.obtido} "
        f"(dif {d.diferenca})"
        for d in rel.divergencias
    ]
    linhas += [f"faltando {x}" for x in rel.faltando] + [f"sobrando {x}" for x in rel.sobrando]
    explicadas = [(linha, e) for linha in linhas if (e := explicacao(linha, conhecidas))]
    a_explicar = [linha for linha in linhas if not explicacao(linha, conhecidas)]
    print(f"Comparados: {rel.comparados}")
    print(
        f"Divergências: {len(rel.divergencias)} | faltando: {len(rel.faltando)} "
        f"| sobrando: {len(rel.sobrando)} | explicadas: {len(explicadas)} "
        f"| a explicar: {len(a_explicar)}"
    )
    for linha in a_explicar[:50]:
        print(f"  {linha}")
    if relatorio:
        texto = ["# A explicar", *a_explicar, "", "# Explicadas"]
        texto += [f"{linha} -> {e.motivo}" for linha, e in explicadas]
        relatorio.write_text("\n".join(texto) + "\n", encoding="utf-8")
        print(f"Relatório completo em {relatorio}")
    if a_explicar:
        raise typer.Exit(1)


notas_app = typer.Typer(help="Notas de corretagem (PDF SINACOR).", no_args_is_help=True)
app.add_typer(notas_app, name="notas")


@notas_app.command("extrair")
def notas_extrair(
    origem: Path | None = typer.Option(None, help="Pasta dos PDFs (padrão: _notas-corretagens)."),
    destino: Path | None = typer.Option(
        None, help="Pasta do candidato (padrão: gabarito-candidato, na pasta de dados)."
    ),
    gabarito: Path | None = OpcaoCaminho,
    explicacoes: Path | None = typer.Option(
        None, help="CSV de divergências conhecidas (padrão: na pasta do gabarito)."
    ),
    do_banco: bool = typer.Option(
        False, help="Usa as notas gravadas no banco em vez de extrair os PDFs."
    ),
    banco: Path | None = OpcaoBanco,
) -> None:
    """Extrai e concilia todas as notas e grava um candidato no formato do gabarito.

    Sai com código 1 se alguma nota for rejeitada sem estar no CSV de divergências conhecidas.
    """
    from coir.erros import ExtracaoError
    from coir.gabarito import candidato as cand
    from coir.gabarito.comparar import carregar_explicacoes
    from coir.gabarito.exportar import ARQUIVO_MAPA, PASTA_PRIVADA, exigir_fora_do_repositorio
    from coir.gabarito.pseudonimos import MapaPseudonimos
    from coir.parsing.extrair import Rejeitada, extrair_pasta
    from coir.parsing.normalizar import classes_das_series, operacoes

    dados = dados_dir(exigir=True)
    pasta_pdf = origem or dados / "_notas-corretagens"
    saida = destino or dados / "gabarito-candidato"
    exigir_fora_do_repositorio(saida)
    if do_banco:
        from coir.db import repositorio as repo
        from coir.db.sessao import conectar
        from coir.parsing.extrair import ResultadoExtracao

        res = ResultadoExtracao()
        with conectar(_banco(banco)) as fabrica, fabrica() as sessao:
            normalizadas = repo.carregar_notas(sessao)
    else:
        res = extrair_pasta(pasta_pdf)
        normalizadas = []
    refs = _refs()
    series = classes_das_series(res.notas, refs)
    for nota in res.notas:
        try:
            normalizadas.append((nota, operacoes(nota, refs, series)))
        except ExtracaoError as erro:
            motivo = f"nota {nota.numero}: {erro}"
            res.rejeitadas.append(
                Rejeitada(nota.arquivo, nota.paginas, motivo, nota.numero, nota.data_pregao)
            )
    mapa = MapaPseudonimos.carregar(_destino(gabarito) / PASTA_PRIVADA / ARQUIVO_MAPA)
    candidato = cand.montar(normalizadas, mapa)
    cand.gravar(candidato, mapa, saida)
    md = [
        "# Extração das notas",
        "",
        f"- Notas extraídas e conciliadas: {len(normalizadas)}",
        f"- Operações: {len(candidato.operacoes)}",
        f"- Rejeitadas: {len(res.rejeitadas)}",
        f"- Avisos: {len(res.avisos) + sum(len(n.avisos) for n, _ in normalizadas)}",
        "",
        "## Rejeitadas",
        "",
        *(f"- {r.arquivo} p. {list(r.paginas)}: {r.motivo}" for r in res.rejeitadas),
        "",
        "## Avisos",
        "",
        *(f"- {a}" for a in res.avisos),
        *(f"- {n.arquivo} nota {n.numero}: {a}" for n, _ in normalizadas for a in n.avisos),
        "",
    ]
    (saida / "extracao.md").write_text("\n".join(md), encoding="utf-8")
    print("\n".join(md[2:6]))
    arquivo_exp = explicacoes or _destino(gabarito) / "divergencias-conhecidas.csv"
    conhecidas = {(e.nota, e.data) for e in carregar_explicacoes(arquivo_exp)}
    novas = [r for r in res.rejeitadas if (r.nota, r.data) not in conhecidas]
    for r in res.rejeitadas:
        marca = "" if r in novas else " (conhecida)"
        print(f"  rejeitada{marca}: {r.arquivo} p. {list(r.paginas)}: {r.motivo}")
    for aviso in res.avisos:
        print(f"  aviso: {aviso}")
    print(f"Candidato em {saida}")
    if novas:
        raise typer.Exit(1)


# ------------------------------------------------------------------------------- banco (F3)
banco_app = typer.Typer(
    help="Banco de dados local (coir.db na pasta de dados).", no_args_is_help=True
)
app.add_typer(banco_app, name="banco")
titular_app = typer.Typer(help="Titulares (CPF) do banco.", no_args_is_help=True)
app.add_typer(titular_app, name="titular")


def _banco(banco: Path | None) -> Path:
    from coir.db.sessao import NOME_BANCO

    return banco or dados_dir(exigir=True) / NOME_BANCO


def _agora() -> "datetime":
    from datetime import datetime

    return datetime.now().replace(microsecond=0)


@banco_app.command("migrar")
def banco_migrar(banco: Path | None = OpcaoBanco) -> None:
    """Cria o banco ou o leva à última versão do esquema (backup antes)."""
    from coir.db.sessao import migrar

    caminho = _banco(banco)
    migrar(caminho, _agora())
    print(f"Banco atualizado: {caminho}")


@titular_app.command("adicionar")
def titular_adicionar(
    cpf: str = typer.Argument(..., help="CPF do titular (com ou sem máscara)."),
    nome: str = typer.Option(..., help="Nome do titular."),
    banco: Path | None = OpcaoBanco,
) -> None:
    """Cadastra um titular (CPF) no banco."""
    from coir.db import repositorio as repo
    from coir.db.sessao import conectar, migrar

    caminho = _banco(banco)
    migrar(caminho, _agora())
    with conectar(caminho) as fabrica, fabrica.begin() as sessao:
        titular = repo.adicionar_titular(sessao, cpf, nome)
        print(f"Titular cadastrado: {titular.nome} (id {titular.id})")


@titular_app.command("listar")
def titular_listar(banco: Path | None = OpcaoBanco) -> None:
    """Lista os titulares do banco."""
    from coir.db import repositorio as repo
    from coir.db.sessao import conectar

    with conectar(_banco(banco)) as fabrica, fabrica() as sessao:
        for t in repo.titulares(sessao):
            print(f"{t.id}: {t.nome}")


@titular_app.command("renomear")
def titular_renomear(
    id_titular: int = typer.Argument(..., help="Id do titular (coir titular listar)."),
    nome: str = typer.Option(..., help="Novo nome."),
    banco: Path | None = OpcaoBanco,
) -> None:
    """Troca o nome de um titular (o criado na importação usa o nome impresso na nota)."""
    from coir.db import repositorio as repo
    from coir.db.sessao import conectar

    with conectar(_banco(banco)) as fabrica, fabrica.begin() as sessao:
        titular = repo.renomear_titular(sessao, id_titular, nome)
        print(f"Titular {titular.id}: {titular.nome}")


@app.command("importar")
def importar(
    pasta: Path | None = typer.Argument(None, help="Pasta dos PDFs (padrão: _notas-corretagens)."),
    titular: str | None = typer.Option(
        None,
        help="CPF do titular (opcional: cada nota vai para o CPF impresso nela; só confere "
        "e cobre notas sem CPF válido).",
    ),
    banco: Path | None = OpcaoBanco,
    explicacoes: Path | None = typer.Option(
        None, help="CSV de divergências conhecidas (padrão: na pasta do gabarito)."
    ),
) -> None:
    """Importa as notas da pasta para o banco, numa transação e sem duplicar.

    Sai com código 1 se alguma nota for rejeitada sem estar no CSV de divergências conhecidas.
    """
    from coir.gabarito.comparar import carregar_explicacoes
    from coir.servicos.importacao import importar_pasta

    dados = dados_dir(exigir=True)
    rel = importar_pasta(
        _banco(banco), pasta or dados / "_notas-corretagens", _refs(), _agora(),
        titular,
    )  # fmt: skip
    print(
        f"Arquivos lidos: {len(rel.arquivos_lidos)}; já importados: {len(rel.arquivos_ignorados)}"
    )
    print(f"Notas gravadas: {rel.notas_gravadas} ({rel.operacoes_gravadas} operações); "
          f"já no banco: {rel.notas_ja_no_banco}; rejeitadas: {len(rel.rejeitadas)}")  # fmt: skip
    if rel.notas_renormalizadas:
        print(f"Notas já gravadas com operações regravadas: {rel.notas_renormalizadas}")
    for criado in rel.titulares_criados:
        print(f"Titular criado: {criado} (coir titular listar / renomear)")
    arquivo_exp = explicacoes or dados / "gabarito" / "divergencias-conhecidas.csv"
    conhecidas = {(e.nota, e.data) for e in carregar_explicacoes(arquivo_exp)}
    novas = [r for r in rel.rejeitadas if (r.nota, r.data) not in conhecidas]
    for r in rel.rejeitadas:
        marca = "" if r in novas else " (conhecida)"
        print(f"  rejeitada{marca}: {r.arquivo} p. {list(r.paginas)}: {r.motivo}")
    for aviso in rel.avisos:
        print(f"  aviso: {aviso}")
    if novas:
        raise typer.Exit(1)


@gabarito_app.command("exportar-eventos")
def gabarito_exportar_eventos(
    pasta: Path | None = OpcaoCaminho,
    destino: Path | None = typer.Option(None, help="Padrão: eventos.csv na pasta de dados."),
) -> None:
    """Converte os eventos da planilha para o CSV de eventos do COIR (com as contas reais)."""
    from coir.gabarito.eventos import exportar_eventos

    arquivo = destino or dados_dir(exigir=True) / ARQUIVO_EVENTOS
    n = exportar_eventos(_destino(pasta), arquivo)
    print(f"{n} eventos em {arquivo} (revise e acrescente os que a planilha não tinha)")


# ------------------------------------------------------------------------------ eventos (F4)
ARQUIVO_EVENTOS = "eventos.csv"
eventos_app = typer.Typer(help="Eventos corporativos e de custódia.", no_args_is_help=True)
app.add_typer(eventos_app, name="eventos")


@eventos_app.command("importar")
def eventos_importar(
    arquivo: Path | None = typer.Argument(
        None, help="CSV (padrão: eventos.csv na pasta de dados)."
    ),
    banco: Path | None = OpcaoBanco,
    substituir: bool = typer.Option(
        False, help="Remove do banco os eventos que não estão mais no CSV."
    ),
) -> None:
    """Grava no banco os eventos novos do CSV (reimportar não duplica)."""
    from coir.db.sessao import conectar, fazer_backup, migrar
    from coir.servicos.eventos import importar_eventos, ler_csv

    caminho = _banco(banco)
    agora = _agora()
    migrar(caminho, agora)
    linhas = ler_csv(arquivo or dados_dir(exigir=True) / ARQUIVO_EVENTOS)
    fazer_backup(caminho, agora)  # regra 8: backup antes de importar
    with conectar(caminho) as fabrica, fabrica() as sessao:
        rel = importar_eventos(sessao, linhas, substituir)
    print(f"Eventos novos: {rel.novos}; já no banco: {rel.ja_no_banco}; removidos: {rel.removidos}")
    if rel.fora_do_csv:
        print(
            f"Aviso: {rel.fora_do_csv} evento(s) no banco que não estão no CSV "
            "(linha corrigida ou apagada); use --substituir para o banco seguir o CSV"
        )


b3_app = typer.Typer(help="Planilhas da Área do Investidor da B3.", no_args_is_help=True)
app.add_typer(b3_app, name="b3")
PASTA_B3 = "_relatorios-b3"


@b3_app.command("importar")
def b3_importar(
    pasta: Path | None = typer.Argument(None, help="Pasta dos xlsx (padrão: _relatorios-b3)."),
    titular: int | None = typer.Option(None, help="Id do titular (coir titular listar)."),
    banco: Path | None = OpcaoBanco,
) -> None:
    """Grava movimentação, negociação e posição (consolidado) da B3; reimportar não duplica."""
    from coir.db import repositorio as repo
    from coir.db.sessao import conectar, migrar
    from coir.servicos.b3 import importar_planilhas

    caminho, agora = _banco(banco), _agora()
    origem = pasta or dados_dir(exigir=True) / PASTA_B3
    arquivos = sorted(p for p in origem.glob("*.xlsx") if not p.name.startswith("~$"))
    migrar(caminho, agora)
    if titular is None:
        with conectar(caminho) as fabrica, fabrica() as s:
            ids = [t.id for t in repo.titulares(s)]
        if len(ids) != 1:
            raise typer.BadParameter(f"escolha o titular com --titular ({ids or 'nenhum'})")
        titular = ids[0]
    rel = importar_planilhas(caminho, arquivos, titular, agora)
    for a in rel.arquivos:
        tipo = a.tipo.value if a.tipo else "ignorado"
        print(f"  {a.arquivo}: {tipo}, {a.novas} nova(s) de {a.linhas} ({a.detalhe})")
    print(f"Linhas novas: {rel.novas}")


# ------------------------------------------------------------------------------ apuração (F4)
def _mes(texto: str) -> tuple[int, int]:
    import re

    m = re.fullmatch(r"(\d{4})-(\d{2})", texto)
    if m is None or not 1 <= int(m[2]) <= 12:
        raise typer.BadParameter("use AAAA-MM")
    return int(m[1]), int(m[2])


@app.command("apurar")
def apurar(
    mes: str | None = typer.Argument(None, help="Mês AAAA-MM (padrão: o último apurado)."),
    ano: int | None = typer.Option(None, help="Resumo dos 12 meses do ano."),
    banco: Path | None = OpcaoBanco,
) -> None:
    """Apuração mensal (comum, day trade, FII), IRRF e DARF, com a memória de cálculo do mês."""
    from datetime import date

    from coir.servicos.apuracao import apurar_banco
    from coir.servicos.relatorio_apuracao import resumo_ano, resumo_mes

    agora = _agora().date()
    fiscais = _fiscais()
    ap = apurar_banco(_banco(banco), _refs(), fiscais, agora)
    if ano is not None and mes is not None:
        raise typer.BadParameter("use AAAA-MM ou --ano, não os dois")
    if ano is not None:
        print(resumo_ano(ap, ano))
        return
    alvo = _mes(mes) if mes else (agora.year, agora.month)
    codigo = fiscais.parametros.texto("darf_codigo", date(alvo[0], alvo[1], 1))
    print(resumo_mes(ap, *alvo, codigo))
    if alvo == (agora.year, agora.month):
        print("(mês em andamento: apuração parcial)")


@app.command("carteira")
def carteira(
    data: str | None = typer.Argument(None, help="Data AAAA-MM-DD (padrão: hoje)."),
    visao: str = typer.Option(
        "titular",
        help="titular (PM fiscal: apuração e IRPF), corretora ou conta (visões operacionais).",
    ),
    banco: Path | None = OpcaoBanco,
) -> None:
    """Posições abertas e PM no fim do dia, na visão pedida."""
    from datetime import date

    from coir.domain.fiscal import Visao
    from coir.servicos.apuracao import carteira_banco
    from coir.servicos.relatorio_apuracao import carteira_texto

    try:
        alvo = date.fromisoformat(data) if data else _agora().date()
        vista = Visao(visao)
    except ValueError as erro:
        raise typer.BadParameter(str(erro)) from erro
    posicoes, contas = carteira_banco(_banco(banco), _refs(), _fiscais(), alvo, vista)
    print(carteira_texto(posicoes, contas, alvo, vista))


@app.command("exportar")
def exportar(
    ano: int = typer.Option(..., help="Ano-calendário."),
    titular: int | None = typer.Option(None, help="Id do titular (coir titular listar)."),
    formato: str = typer.Option("xlsx,csv,pdf", help="Formatos separados por vírgula."),
    relatorio: str | None = typer.Option(
        None, help="Relatórios separados por vírgula (padrão: todos; veja a lista no erro)."
    ),
    data: str | None = typer.Option(
        None, help="Data da carteira AAAA-MM-DD (padrão: 31/12 do ano ou hoje)."
    ),
    visao: str = typer.Option("titular", help="Visão da carteira: titular, corretora ou conta."),
    banco: Path | None = OpcaoBanco,
    destino: Path | None = typer.Option(
        None, help="Pasta de saída (padrão: relatorios/<ano> na pasta de dados)."
    ),
) -> None:
    """Exporta os relatórios fiscais do ano em xlsx, csv e pdf (sem precisar do Excel)."""
    from datetime import date
    from decimal import Decimal

    from coir.domain.fiscal import Visao
    from coir.servicos.relatorios import dados_banco, validar
    from coir.servicos.relatorios import exportar as gravar

    try:
        dia = date.fromisoformat(data) if data else None
        vista = Visao(visao)
    except ValueError as erro:
        raise typer.BadParameter(str(erro)) from erro
    pasta = destino or dados_dir(exigir=True) / "relatorios" / str(ano)
    formatos = tuple(f.strip().lower() for f in formato.split(",") if f.strip())
    chaves = tuple(c.strip() for c in relatorio.split(",")) if relatorio else None
    validar(pasta, formatos, chaves)
    from coir import referencias

    dados = dados_banco(_banco(banco), _refs(), _fiscais(), ano, _agora(), titular, dia, vista,
                        referencias.pasta_efetiva(dados_dir()))  # fmt: skip
    arquivos = gravar(dados, pasta, formatos, chaves)
    meses = [m for m in dados.apuracao.meses if m.titular == dados.titular and m.ano == ano]
    total = sum((m.darf for m in meses), Decimal(0))
    print(f"Titular {dados.titular}, ano {ano}: {len(arquivos)} arquivo(s) em {pasta}")
    for a in arquivos:
        print(f"  {a.name}")
    print(f"DARF do ano: {total:.2f} em {sum(1 for m in meses if m.darf)} mês(es)")


@app.command("ui")
def abrir_ui(
    banco: Path | None = OpcaoBanco,
    navegador: bool = typer.Option(False, help="Abre no navegador em vez da janela própria."),
    porta: int | None = typer.Option(None, help="Porta local (padrão: uma livre)."),
    abrir: bool = typer.Option(True, help="Abre a janela/navegador ao iniciar."),
) -> None:
    """Abre a interface do COIR (janela própria; sem WebView2, o navegador padrão)."""
    from coir.db.sessao import NOME_BANCO
    from coir.servicos.estado import Estado
    from coir.ui.janela import iniciar

    iniciar(Estado(banco or dados_dir() / NOME_BANCO), navegador=navegador, porta=porta,
            abrir=abrir)  # fmt: skip


def _refs() -> "Referencias":
    """Tabelas de referência: o pacote baixado em <dados>/referencias, se mais novo."""
    from coir import referencias

    return referencias.carregar(referencias.pasta_efetiva(dados_dir()))


def _fiscais() -> "ReferenciasFiscais":
    from coir import referencias

    return referencias.carregar_fiscais(referencias.pasta_efetiva(dados_dir()))


def main() -> None:
    """Ponto de entrada: erros de domínio viram mensagem curta e código de saída 1."""
    try:
        app()
    except CoirError as erro:
        print(f"Erro: {erro}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
