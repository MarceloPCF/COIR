"""Uso:
  uv run python -m tools.referencias atualizar [--data AAAA-MM-DD] [--pasta data] [--resumo r.md]
  uv run python -m tools.referencias pacote --saida dist/   (zip de data/ + manifesto, no release)

``atualizar`` coleta as fontes públicas, regrava em data/ os arquivos gerados que mudaram e escreve
o resumo das mudanças (corpo do PR semanal). Sai com código 0 mesmo sem mudança.
"""

import argparse
import logging
import sys
from datetime import date
from pathlib import Path

from tools.referencias import fontes, gerar, pacote

DATA = Path(__file__).resolve().parents[2] / "data"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="python -m tools.referencias")
    sub = parser.add_subparsers(dest="comando", required=True)
    atualizar = sub.add_parser("atualizar", help="coleta as fontes e regrava data/")
    atualizar.add_argument("--data", type=date.fromisoformat, default=date.today())
    atualizar.add_argument("--pasta", type=Path, default=DATA)
    atualizar.add_argument("--resumo", type=Path, default=None)
    atualizar.add_argument("--bruto", type=Path, default=None, help="salva as respostas aqui")
    atualizar.add_argument("--de-bruto", type=Path, default=None, help="reprocessa sem rede")
    montar = sub.add_parser("pacote", help="zip de data/ com manifesto (SHA-256)")
    montar.add_argument("--pasta", type=Path, default=DATA)
    montar.add_argument("--saida", type=Path, required=True)
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    if args.comando == "atualizar":
        if args.de_bruto is not None:
            coleta = gerar.Coleta.carregar(args.de_bruto)
        else:
            coleta = fontes.coletar(args.data)
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
    zip_ = pacote.montar(args.pasta, args.saida)
    print(zip_)
    return 0


if __name__ == "__main__":
    sys.exit(main())
