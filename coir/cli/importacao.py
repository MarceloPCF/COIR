"""Comandos de importação: notas, eventos e planilhas da B3."""

from pathlib import Path
from typing import TYPE_CHECKING

import typer

from coir.config import dados_dir

if TYPE_CHECKING:
    pass
from coir.cli.comum import ARQUIVO_EVENTOS, OpcaoBanco, _agora, _banco, _refs, app


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
    senha: list[str] | None = typer.Option(
        None,
        help="Senha dos PDFs protegidos (repita a opção para várias); o começo do CPF dos "
        "titulares cadastrados é tentado sem precisar informar.",
    ),
    reprocessar: bool = typer.Option(
        False,
        "--reprocessar",
        help="Relê também as notas já importadas: a releitura vale sobre o que está gravado. "
        "Faz backup antes; nota que a releitura não confirma continua gravada.",
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
        titular, senhas=senha or (), reprocessar=reprocessar,
    )  # fmt: skip
    print(
        f"Arquivos lidos: {len(rel.arquivos_lidos)}; já importados: {len(rel.arquivos_ignorados)}"
    )
    print(f"Notas gravadas: {rel.notas_gravadas} ({rel.operacoes_gravadas} operações); "
          f"já no banco: {rel.notas_ja_no_banco}; rejeitadas: {len(rel.rejeitadas)}")  # fmt: skip
    if rel.notas_renormalizadas:
        print(f"Notas já gravadas com operações regravadas: {rel.notas_renormalizadas}")
    if reprocessar:
        print(f"Reprocessamento: {rel.notas_substituidas} substituída(s), "
              f"{rel.notas_removidas} removida(s)")  # fmt: skip
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
