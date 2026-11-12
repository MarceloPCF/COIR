"""Gera docs/fontes-fiscais.md: cada norma citada no COIR e onde ela é usada.

Uso:
  uv run python -m tools.fontes_fiscais              (regrava o documento)
  uv run python -m tools.fontes_fiscais --verificar  (sai com 1 se o documento estiver velho)

Fontes lidas (as citações já estão nelas; nada é digitado à mão aqui):
- ``data/*.csv`` com coluna ``fonte`` (parâmetros fiscais, isenção de PME, vencimentos);
- ``data/irpf/*.yaml`` (mapeamento das fichas do IRPF por exercício), chave ``fonte``;
- a tabela R1 a R18 e as seções de ``docs/regras-fiscais.md``;
- as docstrings dos casos fiscais em ``tests/fiscal_cases`` e dos módulos de ``coir/engine``;
- ``docs/fontes-dados.md`` (as fontes públicas do pacote de referências), copiado como está.

Determinístico: mesma entrada, mesmos bytes (o teste confere que o documento está em dia).
"""

from __future__ import annotations

import ast
import csv
import re
import sys
from collections import defaultdict
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path

import yaml

RAIZ = Path(__file__).resolve().parents[1]
DESTINO = RAIZ / "docs" / "fontes-fiscais.md"
FONTES_DADOS = "docs/fontes-dados.md"  # fontes do pacote de referências (dados, não normas)

# Citação -> nome normalizado da norma. A ordem importa: o primeiro padrão que casa num trecho
# fica com ele (ex.: "Lei Complementar" antes de "Lei").
PADROES: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("Lei Complementar", re.compile(r"\b(?:LC|Lei Complementar)\s+(\d{1,3})/(\d{4})")),
    ("Lei", re.compile(r"\bLei\s+(\d{1,2}\.\d{3}|\d{1,4})(?:/(\d{4}))?")),
    ("Medida Provisória", re.compile(r"\bMP\s+(\d{1,4}(?:-\d+)?)/(\d{4})")),
    ("Instrução Normativa RFB", re.compile(r"\bIN\s+RFB\s+(\d{1,2}\.?\d{3})/(\d{4})")),
    ("Instrução Normativa SRF", re.compile(r"\bIN\s+SRF\s+(\d{1,4})/(\d{4})")),
    ("Solução de Consulta Cosit", re.compile(r"\bSC\s+Cosit\s+(\d{1,4})/(\d{4})")),
    ("Resolução CMN", re.compile(r"\bRes(?:olução|\.)\s+CMN\s+(\d\.?\d{3})/(\d{4})")),
    ("RIR", re.compile(r"\bRIR/(\d{4})()")),
    ("Perguntas e Respostas IRPF", re.compile(r"\bP&R\s+IRPF(?:\s+(\d{4}))?()")),
    ("Manual MIR", re.compile(r"\bManual\s+MIR(?:\s+(\d{4}))?()")),
    ("CTN", re.compile(r"\bCTN\b()()")),
    ("Súmula STJ", re.compile(r"\bSúmula\s+(\d{1,4})()")),
    ("REsp STJ", re.compile(r"\bREsp\s+(\d\.\d{3}\.\d{3}/[A-Z]{2})()")),
)
TEXTO_IRPF = "programa IRPF"


@dataclass(frozen=True, order=True)
class Uso:
    arquivo: str  # relativo à raiz do repositório, com /
    onde: str  # parâmetro, linha da tabela, teste ou seção
    trecho: str  # o texto da citação (uma linha)


def normas(texto: str) -> list[str]:
    """Normas citadas num texto, normalizadas ("Lei 11.033/2004", "IN RFB 1.585/2015")."""
    achadas: list[str] = []
    ocupado: list[tuple[int, int]] = []
    for tipo, padrao in PADROES:
        for m in padrao.finditer(texto):
            if any(a <= m.start() < b for a, b in ocupado):
                continue
            ocupado.append(m.span())
            numero, ano = m.group(1) or "", m.group(2) or ""
            nome = _nome(tipo, numero, ano)
            if nome not in achadas:
                achadas.append(nome)
    if TEXTO_IRPF in texto and "Programa IRPF" not in achadas:
        achadas.append("Programa IRPF")
    return achadas


def _nome(tipo: str, numero: str, ano: str) -> str:
    curto = {"Lei Complementar": "LC", "Medida Provisória": "MP", "Instrução Normativa RFB":
             "IN RFB", "Solução de Consulta Cosit": "SC Cosit", "Resolução CMN": "Res. CMN",
             "Súmula STJ": "Súmula STJ", "REsp STJ": "REsp"}.get(tipo, tipo)  # fmt: skip
    if tipo == "RIR":
        return f"RIR/{numero}"
    if tipo in ("Perguntas e Respostas IRPF", "Manual MIR"):
        return f"{'P&R IRPF' if tipo.startswith('Perguntas') else 'Manual MIR'}" + (
            f" {numero}" if numero else "")  # fmt: skip
    if tipo == "CTN":
        return "CTN"
    return f"{curto} {numero}" + (f"/{ano}" if ano else "")


def _linha(texto: str) -> str:
    return " ".join(texto.split())


# Anotações de desenvolvimento (quem decidiu, quando, em que fase) ficam no código e em
# docs/regras-fiscais.md; no documento publicado viram "critério do COIR". A ordem importa.
_DATA = r"\d{1,2}(?:-\d{1,2})?/\d{1,2}(?:/\d{4})?"  # 19/10/2026, 14-15/10/2026, 20/10
_DATAS = rf"{_DATA}(?:(?:, | e ){_DATA})*"
# datas do desenvolvimento (set. a dez./2026): só elas somem quando aparecem soltas entre
# parênteses; "(19/11/2025)" ou "(01/01/2026)" citam versão de página ou vigência e ficam
_DEV = r"\d{1,2}(?:-\d{1,2})?/(?:0?9|1[0-2])/2026"
_DEVS = rf"{_DEV}(?:(?:, | e ){_DEV})*"
_ESPECIALISTA = r"(?:do especialista(?: fiscal)?|do especialista-fiscal-b3)"
LIMPEZA: tuple[tuple[re.Pattern[str], str], ...] = tuple(
    (re.compile(padrao, re.IGNORECASE), troca)
    for padrao, troca in (
        # "parecer do especialista de 09/11/2026, seguido pelo usuário", "pareceres de 20 e 21/10"
        (rf"pareceres(?: {_ESPECIALISTA})?(?: de)? (?:\d{{1,2}} e )?{_DATAS}", "critérios do COIR"),
        (
            rf"parecer(?: {_ESPECIALISTA})?(?:,| de)? {_DATAS},? seguido pelo usuário",
            "critério do COIR",
        ),
        (rf"parecer(?: {_ESPECIALISTA})?(?:,| de)? {_DATAS}", "critério do COIR"),
        (rf"parecer {_ESPECIALISTA}", "critério do COIR"),
        # "decisão do usuário de 03/11/2026, com parecer…", "decisões da F1 e do usuário, 15/10"
        (rf"decisões da F\d+ e do usuário(?:, {_DATAS})?", "critérios do COIR"),
        (rf"decisões do usuário \({_DATAS}\)", "critérios do COIR"),
        (rf"decisões do usuário(?: de)? {_DATAS}", "critérios do COIR"),
        (rf"decisão do usuário \({_DATAS}\)", "critério do COIR"),
        (rf"decisão do usuário(?: de)? {_DATAS}(?:, com critério do COIR)?", "critério do COIR"),
        (r"decisões do usuário", "critérios do COIR"),
        (r"decisão do usuário", "critério do COIR"),
        (rf"(?:desde )?a decisão de {_DATAS}", "o critério do COIR"),
        (rf"decisão de {_DATAS}", "critério do COIR"),
        (
            rf"(?:escolhida|confirmad[oa]|conferid[oa]) pelo usuário(?: em {_DATAS})?",
            "adotada no COIR",
        ),
        (r",? seguido pelo usuário", ""),
        (r"\bpelo (?:usuário|mantenedor)\b", "no COIR"),
        (r"\bleitura A\d+,? ", ""),
        # fases do projeto: "(F8, 20/10/2026)", "da F4 (motor fiscal, 14-15/10/2026)", "na F8"
        (r" ?\(F\d+[a-z]?(?:, [^()]*?)?\)", ""),
        (rf" ?\((?:[^()]*?, )?{_DEVS}\)(?=[:.;,]| [A-Z|]|$)", ""),
        (r"\b(?:da|na|pela|desde a) F\d+[a-z]?\b,? ?", ""),
        (rf"\bF\d+[a-z]?, {_DATAS}", ""),
        (rf"decisão delegada ao especialista em {_DATAS}", "critério do COIR"),
        (r" \(não foi objeto do parecer\)", ""),
        (r" \(aba [^()]* do v2\)", ""),
        # títulos de seção de docs/regras-fiscais.md e sobras do corte do trecho
        (r"^Decisões(?: d[ao] [^:]+)?\s*$", "Critérios do COIR"),
        (r"Casos do especialista fiscal, conferidos à mão", "Casos conferidos à mão"),
        (r"Casos do especialista fiscal", "Casos conferidos à mão"),
        (r"Casos fechados à mão pelo especialista", "Casos fechados à mão"),
        (rf"^{_DEVS}\)[:*]*\s*", ""),
        (r" ?\(?\b(?:parecer|pareceres|decisões|decisão)(?: [a-z]+){0,2}…$", "…"),
    )
)


def _publico(texto: str) -> str:
    """O trecho como sai no documento publicado, sem as anotações de desenvolvimento."""
    for padrao, troca in LIMPEZA:
        texto = padrao.sub(troca, texto)
    texto = re.sub(r"critério do COIR(?:[,;] critério do COIR)+", "critério do COIR", texto)
    texto = re.sub(r"\(\s*[,;]?\s*\)", "", texto)
    texto = re.sub(r"(^|[.:] )crit(é)rio", r"\1Crit\2rio", texto)
    return re.sub(r" {2,}", " ", texto).strip()


def _rel(caminho: Path) -> str:
    return caminho.relative_to(RAIZ).as_posix()


def _csvs(raiz: Path) -> Iterator[tuple[str, Uso]]:
    for caminho in sorted((raiz / "data").glob("*.csv")):
        with caminho.open(encoding="utf-8", newline="") as f:
            leitor = csv.DictReader(f)
            if "fonte" not in (leitor.fieldnames or ()):
                continue
            chave = (leitor.fieldnames or ["?"])[0]
            for r in leitor:
                texto = r.get("fonte") or ""
                for n in normas(texto):
                    yield n, Uso(_rel(caminho), r.get(chave, ""), _linha(texto))


def _yamls(raiz: Path) -> Iterator[tuple[str, Uso]]:
    for caminho in sorted((raiz / "data" / "irpf").glob("*.yaml")):
        dados = yaml.safe_load(caminho.read_text(encoding="utf-8")) or {}
        yield from _yaml_no(caminho, dados, "")


def _yaml_no(caminho: Path, no: object, onde: str) -> Iterator[tuple[str, Uso]]:
    if isinstance(no, dict):
        for k, v in no.items():  # pyright: ignore[reportUnknownVariableType]
            if k == "fonte" and isinstance(v, str):
                for n in normas(v):
                    yield n, Uso(_rel(caminho), onde or "(raiz)", _linha(v))
            else:
                yield from _yaml_no(caminho, v, f"{onde}.{k}" if onde else str(k))
    elif isinstance(no, list):
        for i, v in enumerate(no):  # pyright: ignore[reportUnknownVariableType, reportUnknownArgumentType]
            rotulo = v.get("codigo") or v.get("chave") or i if isinstance(v, dict) else i  # pyright: ignore[reportUnknownMemberType]
            yield from _yaml_no(caminho, v, f"{onde}[{rotulo}]")


def _regras(raiz: Path) -> Iterator[tuple[str, Uso]]:
    caminho = raiz / "docs" / "regras-fiscais.md"
    secao = ""
    for linha in caminho.read_text(encoding="utf-8").splitlines():
        if linha.startswith("#"):
            secao = linha.lstrip("# ").strip()
            continue
        m = re.match(r"\|\s*(R\d+[a-z]?)\s*\|", linha)
        onde = m.group(1) if m else secao
        for n in normas(linha):
            yield n, Uso(_rel(caminho), onde, _janela(linha.strip("|- "), n))


def _docstrings(raiz: Path) -> Iterator[tuple[str, Uso]]:
    pastas = (raiz / "tests" / "fiscal_cases", raiz / "coir" / "engine")
    for caminho in sorted(p for pasta in pastas for p in pasta.glob("*.py")):
        arvore = ast.parse(caminho.read_text(encoding="utf-8"))
        nos: list[tuple[str, ast.AST]] = [("módulo", arvore)]
        nos += [(n.name, n) for n in ast.walk(arvore)
                if isinstance(n, ast.FunctionDef | ast.ClassDef)]  # fmt: skip
        for onde, no in nos:
            doc = ast.get_docstring(no)  # type: ignore[arg-type]
            for n in normas(doc or ""):
                yield n, Uso(_rel(caminho), onde, _janela(doc or "", n))


def _janela(texto: str, norma: str, largura: int = 110) -> str:
    """O trecho em volta da citação (o texto inteiro, se for curto), cortado em palavras."""
    texto = _publico(_linha(texto))  # limpa antes de cortar
    if len(texto) <= 2 * largura:
        return texto
    i = max(texto.find(norma.split("/")[0].split(" ")[-1]), 0)
    inicio, fim = max(i - largura, 0), min(i + largura, len(texto))
    inicio = texto.find(" ", inicio) + 1 if inicio else 0
    fim = texto.rfind(" ", 0, fim) if fim < len(texto) else fim
    return ("…" if inicio else "") + texto[inicio:fim] + ("…" if fim < len(texto) else "")


ORDEM = ("LC", "Lei", "MP", "IN RFB", "Res. CMN", "SC Cosit", "RIR", "P&R IRPF", "Manual MIR",
         "Programa IRPF", "CTN", "Súmula STJ", "REsp")  # fmt: skip


def _ordem(norma: str) -> tuple[int, int, str]:
    """Por tipo de norma e, dentro do tipo, pelo número (leis em ordem cronológica)."""
    tipo = next((i for i, t in enumerate(ORDEM) if norma == t or norma.startswith(t + " ")),
                len(ORDEM))  # fmt: skip
    m = re.search(r" (\d[\d.]*)", norma)
    return tipo, int(m.group(1).replace(".", "")) if m else 0, norma


def _com_ano(usos: dict[str, set[Uso]]) -> dict[str, set[Uso]]:
    """ "Lei 11.033" citada sem ano junta-se a "Lei 11.033/2004" quando o ano aparece em outra
    citação (só se houver um ano possível)."""
    anos: dict[str, set[str]] = defaultdict(set)
    for n in usos:
        if "/" in n and not n.startswith(("RIR", "REsp")):
            anos[n.split("/")[0]].add(n)
    saida: dict[str, set[Uso]] = defaultdict(set)
    for n, u in usos.items():
        alvo = anos.get(n, set())
        saida[next(iter(alvo)) if "/" not in n and len(alvo) == 1 else n] |= u
    return saida


def _ficha(norma: str, raiz: Path) -> str:
    """Arquivo da ficha da norma em docs/normas/ (base normativa), ou "" se não houver."""
    especiais = {"CTN": "lei-5172-1966-ctn", "RIR/2018": "decreto-9580-2018-rir"}
    if norma in especiais:
        nome = especiais[norma]
    elif norma.startswith("P&R IRPF"):
        nome = "pr-irpf"
    elif norma.startswith("Manual MIR"):
        nome = "manual-mir"
    else:
        m = re.match(r"(.+?)\s+([\d.]+)/(\d{4})$", norma)
        if not m:
            return ""
        sigla = m.group(1).replace("Instrução Normativa", "IN").replace("Res.", "Res")
        nome = f"{sigla} {m.group(2).replace('.', '')} {m.group(3)}".lower().replace(" ", "-")
    return f"{nome}.md" if (raiz / "docs" / "normas" / f"{nome}.md").is_file() else ""


def gerar(raiz: Path = RAIZ) -> str:
    brutos: dict[str, set[Uso]] = defaultdict(set)
    for origem in (_csvs, _yamls, _regras, _docstrings):
        for norma, uso in origem(raiz):
            brutos[norma].add(uso)
    usos = _com_ano(brutos)
    partes = [
        "# Fontes fiscais do COIR",
        "",
        "Gerado por `uv run python -m tools.fontes_fiscais` a partir das tabelas de `data/`, de",
        "`docs/regras-fiscais.md` e das docstrings dos casos fiscais e do motor. Não edite à mão:",
        "corrija a citação na origem e gere de novo. O documento orienta; não substitui a norma.",
        "O que cada norma diz (vigência, artigos, pontos em aberto) está na base normativa:",
        "[`docs/normas/`](normas/README.md).",
        "",
        "| Norma | Usos | Ficha |",
        "|---|---|---|",
    ]
    for n in sorted(usos, key=_ordem):
        f = _ficha(n, raiz)
        partes.append(f"| {n} | {len(usos[n])} | {f'[{f}](normas/{f})' if f else '—'} |")
    dados = raiz / FONTES_DADOS
    if dados.is_file():  # copiado como está (edite lá)
        partes += ["", *dados.read_text(encoding="utf-8").rstrip().splitlines()]
    for norma in sorted(usos, key=_ordem):
        partes += ["", f"## {norma}", ""]
        for u in sorted(usos[norma]):
            partes.append(f"- `{u.arquivo}` — {_publico(u.onde)}: {_publico(u.trecho)}")
    return "\n".join(partes) + "\n"


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    texto = gerar()
    if "--verificar" in args:
        atual = DESTINO.read_text(encoding="utf-8") if DESTINO.is_file() else ""
        if atual != texto:
            print("docs/fontes-fiscais.md desatualizado: uv run python -m tools.fontes_fiscais")
            return 1
        return 0
    DESTINO.write_bytes(texto.encode("utf-8"))
    print(f"{DESTINO.relative_to(RAIZ).as_posix()}: {texto.count(chr(10) + '## ')} normas")
    return 0


if __name__ == "__main__":
    sys.exit(main())
