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

OpcaoCaminho = typer.Option(None, help="Padrão: dentro da pasta de dados (COIR_DADOS).")


def _origem(origem: Path | None) -> Path:
    return origem or dados_dir(exigir=True) / "_gabarito" / "gabarito.xlsb"


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
) -> None:
    """Gera piloto.md: gabarito x resumo do PDF para poucas notas."""
    from coir.gabarito import piloto
    from coir.gabarito.exportar import ARQUIVO_MAPA, PASTA_PRIVADA, exigir_fora_do_repositorio

    pasta = _destino(destino)
    exigir_fora_do_repositorio(pasta)
    gabarito, mapa = _ler_planilha(_origem(origem), pasta)
    mapa.salvar(pasta / PASTA_PRIVADA / ARQUIVO_MAPA)
    nomes = tuple(n.strip() for n in notas.split(",") if n.strip()) or piloto.NOTAS_PADRAO
    texto = piloto.gerar(gabarito, dados_dir(exigir=True) / "_notas-corretagens", nomes)
    saida = pasta / "piloto" / "piloto.md"
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(texto, encoding="utf-8")
    print(f"Piloto gravado em {saida} ({len(nomes)} PDF(s)).")


@gabarito_app.command("conferir")
def gabarito_conferir(
    origem: Path | None = OpcaoCaminho, destino: Path | None = OpcaoCaminho
) -> None:
    """Confere automaticamente o gabarito contra todos os PDFs de notas (conferencia.md/.csv)."""
    import csv

    from coir.gabarito import conferencia
    from coir.gabarito.exportar import exigir_fora_do_repositorio

    pasta = _destino(destino)
    exigir_fora_do_repositorio(pasta)
    gabarito, _ = _ler_planilha(_origem(origem), pasta)
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
    titular_unico: bool = typer.Option(
        False, help="Trata todos os CPFs da planilha como um único titular."
    ),
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
    niveis: str = typer.Option("operacoes,posicoes,apuracao", help="Níveis separados por vírgula."),
) -> None:
    """Compara um candidato com o gabarito usando as tolerâncias de data/tolerancias.toml."""
    from coir.gabarito.comparar import Tolerancias, comparar
    from coir.gabarito.exportar import carregar

    tol = Tolerancias.carregar()
    rel = comparar(
        carregar(_destino(gabarito)),
        carregar(candidato),
        tol,
        [n.strip() for n in niveis.split(",") if n.strip()],
    )
    print(f"Comparados: {rel.comparados}")
    print(
        f"Divergências: {len(rel.divergencias)} | faltando: {len(rel.faltando)} "
        f"| sobrando: {len(rel.sobrando)}"
    )
    for d in rel.divergencias[:50]:
        print(
            f"  {d.nivel} {d.chave} {d.campo}: esperado {d.esperado} obtido {d.obtido} "
            f"(dif {d.diferenca})"
        )
    for item in (rel.faltando + rel.sobrando)[:50]:
        print(f"  {item}")
    if not rel.ok:
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
