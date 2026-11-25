"""Comandos da conferência com a planilha-gabarito e a planilha externa (grupo gabarito e validar-
gabarito)."""

from pathlib import Path
from typing import TYPE_CHECKING

import typer

from coir.config import dados_dir

if TYPE_CHECKING:
    from coir.gabarito.modelo import Gabarito
    from coir.gabarito.pseudonimos import MapaPseudonimos
from coir.cli.comum import (
    ARQUIVO_EVENTOS,
    OpcaoBanco,
    OpcaoCaminho,
    OpcaoTitularUnico,
    _agora,
    _banco,
    _destino,
    _fiscais,
    _origem,
    _pasta_refs,
    _refs,
    app,
)

gabarito_app = typer.Typer(
    help="Gabarito: planilha do usuário convertida para validação.", no_args_is_help=True
)
app.add_typer(gabarito_app, name="gabarito")


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


@gabarito_app.command("externo-mensal")
def gabarito_externo_mensal(
    pasta: Path | None = typer.Argument(
        None,
        help="Pasta com os PDFs mensais de outro sistema, um por ano ('*Comuns e Day-Trade*.pdf'; "
        "padrão: _relatorios-b3).",
    ),
    titular: int = typer.Option(1, help="Id do titular (coir titular listar)."),
    banco: Path | None = OpcaoBanco,
    destino: Path | None = OpcaoCaminho,
) -> None:
    """Compara mês a mês a ficha de renda variável de outro sistema com a apuração do motor,
    partindo do prejuízo que ele recebeu no primeiro mês (ex.: o da última DIRPF)."""
    from collections import Counter

    from coir.engine.calendario import fim_do_mes
    from coir.gabarito.comparar import Tolerancias
    from coir.gabarito.exportar import exigir_fora_do_repositorio
    from coir.gabarito.externo_mensal import (
        carregar_explicacoes,
        comparar_mensal,
        explicar,
        gravar,
        ler_relatorios,
        saldo_inicial,
    )
    from coir.servicos.apuracao import apurar_banco

    origem = pasta or dados_dir(exigir=True) / "_relatorios-b3"
    arquivos = sorted(origem.glob("*Comuns e Day-Trade*.pdf"))
    if not arquivos:
        raise typer.BadParameter(f"nenhum '*Comuns e Day-Trade*.pdf' em {origem}")
    meses = ler_relatorios(arquivos)
    inicial = saldo_inicial(meses, titular)
    ultimo = meses[-1]
    apuracao = apurar_banco(_banco(banco), _refs(), _fiscais(),
                            fim_do_mes(ultimo.ano, ultimo.mes), [inicial])  # fmt: skip
    pasta_saida = _destino(destino)
    conhecidas = carregar_explicacoes(pasta_saida / "externo-mensal-explicadas.csv")
    tol = Tolerancias.carregar().mes
    linhas = explicar(comparar_mensal(meses, apuracao.meses, titular, tol), conhecidas)
    print(f"Relatório externo: {len(meses)} meses ({meses[0].mes:02d}/{meses[0].ano} a "
          f"{ultimo.mes:02d}/{ultimo.ano}); prejuízo inicial em 31/12/{inicial.ano}: "
          f"comuns {inicial.comum}, day trade {inicial.day_trade}")  # fmt: skip
    for classe, n in Counter(x.classe for x in linhas).most_common():
        print(f"  {classe}: {n}")
    exigir_fora_do_repositorio(pasta_saida)
    saida = pasta_saida / "externo-mensal-divergencias.csv"
    gravar([x for x in linhas if x.classe != "confere"], saida)
    print(f"Detalhe em {saida}")


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

    from coir.gabarito.comparar import Tolerancias
    from coir.gabarito.exportar import exigir_fora_do_repositorio
    from coir.gabarito.externo import ticker_atual
    from coir.importers.dirpf import ler_dirpf
    from coir.reports.irpf import fichas
    from coir.servicos.declaracao import declaracoes_lidas
    from coir.servicos.dirpf_comparar import carregar_explicacoes, comparar, explicar, gravar
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
    pasta_refs = _pasta_refs(dados)
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
