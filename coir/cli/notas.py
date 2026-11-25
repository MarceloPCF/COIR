"""Comandos das notas de corretagem: extrair e anonimizar."""

from pathlib import Path
from typing import TYPE_CHECKING

import typer

from coir.config import dados_dir

if TYPE_CHECKING:
    pass
from coir.cli.comum import OpcaoBanco, OpcaoCaminho, _banco, _destino, _fiscais, _refs, app

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
        # pasta só com textos de nota (base anonimizada de desenvolvimento): lê os .txt
        textos = not any(pasta_pdf.glob("*.pdf")) and any(pasta_pdf.glob("*.txt"))
        res = extrair_pasta(pasta_pdf, ".txt" if textos else ".pdf")
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
    candidato = cand.montar(normalizadas, mapa, pme=_fiscais().pme)
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


@notas_app.command("anonimizar")
def notas_anonimizar(
    origens: list[Path] = typer.Argument(..., help="PDFs de nota ou pastas com PDFs."),
    saida: Path | None = typer.Option(
        None, help="Pasta das cópias (padrão: suporte/notas-anonimizadas, na pasta de dados)."
    ),
    mapa: Path | None = typer.Option(
        None, help="Mapa de anonimização (padrão: _privado/mapa-anonimizacao.json nos dados)."
    ),
    imagens: bool = typer.Option(
        False, "--imagens", help="Grava também a imagem de cada página, com tarja."
    ),
) -> None:
    """Grava o texto anonimizado de cada nota (formato das fixtures, páginas separadas por \\f).

    Usa os dados para ocultar de cada titular (tela Suporte; nome, CPF e contas do banco). Cada
    cópia passa por dois portões: a extração dela tem de ser igual à do original (só a conta
    muda, sempre pela mesma fictícia) e nenhum dado pessoal do original pode sobrar. Sai com
    código 1 se alguma nota for recusada (a recusa diz a nota e o campo, nunca o valor).
    """
    from coir.db.sessao import NOME_BANCO
    from coir.servicos import suporte

    dados = dados_dir(exigir=True)
    arquivos: list[Path] = []
    for origem in origens:
        arquivos += sorted(origem.glob("*.pdf")) if origem.is_dir() else [origem]
    resultados = suporte.anonimizar_arquivos(
        arquivos,
        saida or dados / suporte.PASTA_SUPORTE / "notas-anonimizadas",
        mapa or dados / suporte.ARQUIVO_MAPA,
        suporte.termos_sensiveis(dados),
        informados=suporte.dados_para_ocultar(dados, dados / NOME_BANCO),
        imagens=imagens,
    )
    recusados = [r for r in resultados if r.destino is None]
    for r in recusados:
        print(f"recusada: {r.origem}")
        for motivo in r.recusas:
            print(f"  - {motivo}")
    gravados = len(resultados) - len(recusados)
    print(f"{gravados} cópia(s) gravada(s), {len(recusados)} recusada(s).")
    if recusados:
        raise typer.Exit(1)
