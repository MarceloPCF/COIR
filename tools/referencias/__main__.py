"""Uso:
  uv run python -m tools.referencias atualizar [--data AAAA-MM-DD] [--pasta data] [--resumo r.md]
  uv run python -m tools.referencias pacote --saida dist/   (zip de data/ + manifesto, no release)

  uv run python -m tools.referencias historico [--cache pasta] [--resumo r.md]
  uv run python -m tools.referencias nomes [--cache pasta] [--resumo r.md]
  uv run python -m tools.referencias pregao [--data AAAA-MM-DD] (0: com pregão; 10: sem pregão)

``atualizar`` coleta as fontes públicas, regrava em data/ os arquivos gerados que mudaram e escreve
o resumo das mudanças (corpo do PR semanal). Sai com código 0 mesmo sem mudança.

``nomes`` refaz ``nomes_pregao_b3.csv`` com o COTAHIST de 2008 em diante (a coleta semanal só lê o
ano corrente). ``pregao`` diz se vale coletar hoje: fim de semana e feriado da B3 ficam para o
próximo dia útil.

``historico`` completa ``eventos_b3.csv`` desde 2008 com o FRE da CVM e o COTAHIST da B3 (uns 1,5 GB
de zips anuais; ``--cache`` guarda e reaproveita). O histórico não muda de uma semana para outra:
roda à mão (ou pelo workflow, sob demanda), e o resumo lista divergências e revisões.
"""

import argparse
import csv
import io
import logging
import sys
from datetime import date
from pathlib import Path

from tools.referencias import fontes, gerar, historico, pacote

DATA = Path(__file__).resolve().parents[2] / "data"
SEM_PREGAO = 10  # código de saída próprio: outro erro não pode passar por feriado


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="python -m tools.referencias")
    sub = parser.add_subparsers(dest="comando", required=True)
    atualizar = sub.add_parser("atualizar", help="coleta as fontes e regrava data/")
    atualizar.add_argument("--data", type=date.fromisoformat, default=date.today())
    atualizar.add_argument("--pasta", type=Path, default=DATA)
    atualizar.add_argument("--resumo", type=Path, default=None)
    atualizar.add_argument("--bruto", type=Path, default=None, help="salva as respostas aqui")
    atualizar.add_argument("--de-bruto", type=Path, default=None, help="reprocessa sem rede")
    hist = sub.add_parser("historico", help="eventos desde 2008 (FRE da CVM + COTAHIST da B3)")
    hist.add_argument("--data", type=date.fromisoformat, default=date.today())
    hist.add_argument("--pasta", type=Path, default=DATA)
    hist.add_argument("--cache", type=Path, default=None, help="zips anuais baixados")
    hist.add_argument("--resumo", type=Path, default=None)
    nomes = sub.add_parser("nomes", help="nomes de pregão desde 2008 (COTAHIST da B3)")
    nomes.add_argument("--data", type=date.fromisoformat, default=date.today())
    nomes.add_argument("--pasta", type=Path, default=DATA)
    nomes.add_argument("--cache", type=Path, default=None, help="zips anuais baixados")
    nomes.add_argument("--resumo", type=Path, default=None)
    dia = sub.add_parser("pregao", help="código 0 com pregão no dia, 10 sem pregão")
    dia.add_argument("--data", type=date.fromisoformat, default=date.today())
    dia.add_argument("--pasta", type=Path, default=DATA)
    montar = sub.add_parser("pacote", help="zip de data/ com manifesto (SHA-256)")
    montar.add_argument("--pasta", type=Path, default=DATA)
    montar.add_argument("--saida", type=Path, required=True)
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    if args.comando == "atualizar":
        if args.de_bruto is not None:
            coleta = gerar.Coleta.carregar(args.de_bruto)
        else:
            coleta = fontes.coletar(args.data, sufixos=fontes.sufixos_especificacao(args.pasta))
            if args.bruto is not None:
                coleta.salvar(args.bruto)
        arquivos, resumo = gerar.gerar(args.pasta, coleta)
        mudaram = gerar.gravar(args.pasta, arquivos)
        texto = resumo.markdown(args.data)
        if args.resumo is not None:
            args.resumo.write_text(texto, encoding="utf-8")
        print(texto)
        print(f"arquivos alterados: {', '.join(mudaram) or 'nenhum'}")
        return 0
    if args.comando == "historico":
        companhias = {r["cnpj"]: r["emissor"]
                      for r in gerar.ler_csv(args.pasta / "emissores_b3.csv")
                      if r["tipo"] == "COMPANHIA"}  # fmt: skip
        pares = historico.pares_de_troca(gerar.ler_csv(args.pasta / "trocas_ticker.csv"))
        eventos, revisao = fontes.historico_eventos(companhias, args.data, args.cache,
                                                    pares=pares)  # fmt: skip
        arquivos, resumo = gerar.incorporar_eventos(args.pasta, eventos, args.data)
        juntos = list(csv.DictReader(io.StringIO(arquivos["eventos_b3.csv"])))
        resumo.revisao = [r.texto() for r in historico.cobertas(revisao, juntos)]
        mudaram = gerar.gravar(args.pasta, arquivos)
        texto = resumo.markdown(args.data)
        if args.resumo is not None:
            args.resumo.write_text(texto, encoding="utf-8")
        print(texto)
        print(f"arquivos alterados: {', '.join(mudaram) or 'nenhum'}")
        return 0
    if args.comando == "nomes":
        sufixos = fontes.sufixos_especificacao(args.pasta)
        resumo = gerar.Resumo()
        historicos = fontes.nomes_historicos(args.data, sufixos, args.cache)
        linhas = gerar.juntar_nomes(gerar.ler_csv(args.pasta / "nomes_pregao_b3.csv"), historicos,
                                    resumo)  # fmt: skip
        mudaram = gerar.gravar(args.pasta, {
            "nomes_pregao_b3.csv": gerar.texto_csv(gerar.COLUNAS_NOMES, linhas),
            "tipos_ativo_b3.csv": gerar.texto_csv(
                ("ticker", "tipo", "observacao"),
                gerar.tipos(gerar.ler_csv(args.pasta / "instrumentos_b3.csv"), linhas,
                            gerar.ler_csv(args.pasta / "emissores_b3.csv")),
            ),
            "pacote.json": (args.pasta / "pacote.json").read_text(encoding="utf-8"),
        })  # fmt: skip
        texto = f"{len(linhas)} nomes; {len(resumo.nomes_novos)} novos"
        if args.resumo is not None:
            args.resumo.write_text(texto, encoding="utf-8")
        print(texto)
        print(f"arquivos alterados: {', '.join(mudaram) or 'nenhum'}")
        return 0
    if args.comando == "pregao":
        tem = gerar.houve_pregao(args.data, gerar.ler_csv(args.pasta / "feriados.csv"))
        print(f"{args.data}: {'com' if tem else 'sem'} pregão")
        return 0 if tem else SEM_PREGAO
    zip_ = pacote.montar(args.pasta, args.saida)
    print(zip_)
    return 0


if __name__ == "__main__":
    sys.exit(main())
