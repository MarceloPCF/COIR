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
    from coir.gabarito.comparar import Tolerancias, carregar_explicacoes, comparar, explicacao
    from coir.gabarito.exportar import carregar

    pasta = _destino(gabarito)
    conhecidas = carregar_explicacoes(explicacoes or pasta / "divergencias-conhecidas.csv")
    rel = comparar(
        carregar(pasta),
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
    from coir import referencias
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
    refs = referencias.carregar()
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


@app.command("importar")
def importar(
    pasta: Path | None = typer.Argument(None, help="Pasta dos PDFs (padrão: _notas-corretagens)."),
    titular: str | None = typer.Option(
        None, help="CPF do titular das contas novas (dispensável com um só titular)."
    ),
    banco: Path | None = OpcaoBanco,
    explicacoes: Path | None = typer.Option(
        None, help="CSV de divergências conhecidas (padrão: na pasta do gabarito)."
    ),
) -> None:
    """Importa as notas da pasta para o banco, numa transação e sem duplicar.

    Sai com código 1 se alguma nota for rejeitada sem estar no CSV de divergências conhecidas.
    """
    from coir import referencias
    from coir.gabarito.comparar import carregar_explicacoes
    from coir.servicos.importacao import importar_pasta

    dados = dados_dir(exigir=True)
    rel = importar_pasta(
        _banco(banco), pasta or dados / "_notas-corretagens", referencias.carregar(), _agora(),
        titular,
    )  # fmt: skip
    print(
        f"Arquivos lidos: {len(rel.arquivos_lidos)}; já importados: {len(rel.arquivos_ignorados)}"
    )
    print(f"Notas gravadas: {rel.notas_gravadas} ({rel.operacoes_gravadas} operações); "
          f"já no banco: {rel.notas_ja_no_banco}; rejeitadas: {len(rel.rejeitadas)}")  # fmt: skip
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


def main() -> None:
    """Ponto de entrada: erros de domínio viram mensagem curta e código de saída 1."""
    try:
        app()
    except CoirError as erro:
        print(f"Erro: {erro}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
