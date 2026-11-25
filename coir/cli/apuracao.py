"""Comandos de resultado: apurar, carteira, exportar e abrir a interface."""

from pathlib import Path
from typing import TYPE_CHECKING

import typer

from coir.config import dados_dir

if TYPE_CHECKING:
    pass
from coir.cli.comum import OpcaoBanco, _agora, _banco, _fiscais, _mes, _pasta_refs, _refs, app


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
    dados = dados_banco(_banco(banco), _refs(), _fiscais(), ano, _agora(), titular, dia, vista,
                        _pasta_refs())  # fmt: skip
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
