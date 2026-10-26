"""Gera data/feriados.csv (dias sem expediente bancário e dias sem pregão na B3).

Uso: uv run python -m tools.gerar_feriados [primeiro_ano] [ultimo_ano]

O CSV é versionado; revise-o ao acrescentar anos (a B3 publica o calendário do ano seguinte em
ofício circular). Tipos:
- bancario: não é dia útil bancário (vale para o DARF e a B3 também fecha): feriados nacionais
  (Lei 662/1949 red. Lei 10.607/2002; Lei 6.802/1980; 20/11 desde 2024, Lei 14.759/2023),
  Sexta-feira Santa, Carnaval (segunda e terça) e Corpus Christi (Res. CMN 4.880/2020 art. 6º,
  antes Res. CMN 2.932/2002 art. 5º) e o último dia útil do ano (Res. 4.880 art. 3º / 2.932
  art. 2º; ex.: DARF de nov/2022 venceu em 29/12/2022);
- pregao: só a B3 fecha: 24/12, os feriados da cidade de São Paulo até 2021 (OC 150/2020-PRE)
  e exceções avulsas (Copa de 2014).
"""

import csv
import sys
from datetime import date, timedelta
from pathlib import Path

DESTINO = Path(__file__).resolve().parents[1] / "data" / "feriados.csv"


def pascoa(ano: int) -> date:
    """Domingo de Páscoa (algoritmo de Meeus/Jones/Butcher, calendário gregoriano)."""
    a, b, c = ano % 19, ano // 100, ano % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    m = (32 + 2 * e + 2 * i - h - k) % 7
    n = (a + 11 * h + 22 * m) // 451
    mes = (h + m - 7 * n + 114) // 31
    dia = (h + m - 7 * n + 114) % 31 + 1
    return date(ano, mes, dia)


def feriados(ano: int) -> list[tuple[date, str, str]]:
    p = pascoa(ano)
    lista = [
        (date(ano, 1, 1), "bancario", "Confraternização Universal"),
        (p - timedelta(days=48), "bancario", "Carnaval (segunda)"),
        (p - timedelta(days=47), "bancario", "Carnaval (terça)"),
        (p - timedelta(days=2), "bancario", "Sexta-feira Santa"),
        (date(ano, 4, 21), "bancario", "Tiradentes"),
        (date(ano, 5, 1), "bancario", "Dia do Trabalho"),
        (p + timedelta(days=60), "bancario", "Corpus Christi"),
        (date(ano, 9, 7), "bancario", "Independência"),
        (date(ano, 10, 12), "bancario", "Nossa Senhora Aparecida"),
        (date(ano, 11, 2), "bancario", "Finados"),
        (date(ano, 11, 15), "bancario", "Proclamação da República"),
        (date(ano, 12, 25), "bancario", "Natal"),
        (date(ano, 12, 24), "pregao", "Véspera de Natal (sem pregão)"),
    ]
    if ano >= 2024:
        lista.append((date(ano, 11, 20), "bancario", "Consciência Negra (Lei 14.759/2023)"))
    elif ano <= 2021:  # B3 abre nos feriados da cidade de São Paulo desde 2022
        lista += [
            (date(ano, 1, 25), "pregao", "Aniversário de São Paulo"),
            (date(ano, 7, 9), "pregao", "Revolução Constitucionalista (SP)"),
            (date(ano, 11, 20), "pregao", "Consciência Negra (SP)"),
        ]
    lista = [x for x in lista if x[0] not in COM_PREGAO]
    lista += [(d, "pregao", n) for d, n in SEM_PREGAO.items() if d.year == ano]
    nao_uteis = {d for d, t, _ in lista if t == "bancario"}
    ultimo = date(ano, 12, 31)
    while ultimo.weekday() >= 5 or ultimo in nao_uteis:
        ultimo -= timedelta(days=1)
    lista.append((ultimo, "bancario", "Último dia útil do ano (sem expediente bancário)"))
    return sorted((d, t, n) for d, t, n in lista if d.weekday() < 5)


# feriados de SP antecipados para maio/2020 (pandemia): a B3 abriu nas datas originais
COM_PREGAO = {date(2020, 7, 9), date(2020, 11, 20)}  # OC 068/2020-PRE
SEM_PREGAO = {date(2014, 6, 12): "Abertura da Copa do Mundo (sem pregão)"}


def main(argv: list[str]) -> None:
    primeiro = int(argv[0]) if argv else 2008
    ultimo = int(argv[1]) if len(argv) > 1 else 2027
    with DESTINO.open("w", encoding="utf-8", newline="") as arquivo:
        escritor = csv.writer(arquivo, lineterminator="\n")
        escritor.writerow(["data", "tipo", "descricao"])
        for ano in range(primeiro, ultimo + 1):
            for dia, tipo, nome in feriados(ano):
                escritor.writerow([dia.isoformat(), tipo, nome])
    print(f"{DESTINO} ({primeiro}-{ultimo})")


if __name__ == "__main__":
    main(sys.argv[1:])
