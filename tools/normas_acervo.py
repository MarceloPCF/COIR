"""Baixa a íntegra das normas da base fiscal (docs/normas/acervo.csv) para uma pasta local.

Uso:
  uv run python -m tools.normas_acervo                 (baixa tudo em .cache/normas/)
  uv run python -m tools.normas_acervo --so IN_RFB_1585_2015 Lei_11033_2004
  uv run python -m tools.normas_acervo --destino PASTA

O acervo é texto oficial copiado, não vai para o Git (``.cache/`` é ignorado). As fichas em
``docs/normas/`` resumem cada norma com as próprias palavras e dizem de onde e quando o texto
foi lido; este script refaz a leitura para a revisão semanal do especialista fiscal.

Fontes (todas oficiais e sem login):
- ``lei``: URN do LexML -> texto compilado da Câmara (Legin, "normaatualizada", com as notas
  "Redação dada pela...") e publicação original (Câmara e Senado). O Planalto recusa conexões
  de linha de comando; o texto compilado dele fica como link na ficha.
- ``sijut``: atos da Receita pela API do SIJUT, visão multivigente (redações tachadas marcadas
  ``[REVOGADO/ANTERIOR]`` e notas de alteração com a data de vigência). Os anexos em PDF (íntegra
  das Soluções de Consulta) são convertidos com ``pdftotext``, se ele existir no PATH.
- ``pdf``: Perguntas e Respostas do IRPF e afins em gov.br (convertidos com ``pdftotext``).
- ``bcb``: normativos do Banco Central (API de conteúdo).
- ``govbr``: páginas do gov.br (Manual MIR, ReVar), com as subpáginas do mesmo caminho.
"""

from __future__ import annotations

import argparse
import csv
import datetime
import html
import json
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.parse
import urllib.request
from collections.abc import Callable
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
LISTA = RAIZ / "docs" / "normas" / "acervo.csv"
DESTINO = RAIZ / ".cache" / "normas"
SIJUT = "https://normasinternet2.receita.fazenda.gov.br/api/consulta-externa/ato"
AGENTE = {"User-Agent": "Mozilla/5.0 (COIR; leitura de normas)"}


def _get(url: str, tempo: int = 120) -> bytes:
    with urllib.request.urlopen(urllib.request.Request(url, headers=AGENTE), timeout=tempo) as r:
        return bytes(r.read())


def _decodifica(b: bytes) -> str:
    try:
        return b.decode("utf-8")
    except UnicodeDecodeError:
        return b.decode("latin-1")


def _texto_html(h: str) -> str:
    h = re.sub(r"(?is)<script.*?</script>|<style.*?</style>", "", h)
    h = re.sub(
        r"(?is)<(strike|s|del)\b[^>]*>(.*?)</\1>",
        lambda m: " [[REVOGADO: " + re.sub(r"<[^>]+>", " ", m.group(2)) + "]] ",
        h,
    )
    h = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</li>|</tr>|</h\d>", "\n", h)
    t = html.unescape(re.sub(r"<[^>]+>", " ", h))
    t = re.sub(r"[ \t\r\xa0]+", " ", t)
    return re.sub(r"\n\s*\n+", "\n", t).strip()


def _pdf_texto(b: bytes) -> str:
    if not shutil.which("pdftotext"):
        return "(pdftotext ausente: PDF não convertido)"
    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "a.pdf"
        p.write_bytes(b)
        r = subprocess.run(
            ["pdftotext", "-enc", "UTF-8", "-layout", str(p), "-"], capture_output=True, check=False
        )
        return r.stdout.decode("utf-8", "replace")


def _cabecalho(fonte: str, **extra: str) -> str:
    linhas = [f"FONTE: {fonte}", f"LIDO EM: {datetime.date.today().isoformat()}"]
    linhas += [f"{k.upper()}: {v}" for k, v in extra.items()]
    return "\n".join(linhas) + "\n----\n"


def lei(nome: str, urn: str, destino: Path) -> list[str]:
    pagina = _decodifica(_get("https://www.lexml.gov.br/urn/" + urn))
    links = [
        html.unescape(x) for x in re.findall(r'href="(http[^"]+)"', pagina) if "linker" not in x
    ]
    camara = [
        x.replace("www2.camara.gov.br", "www2.camara.leg.br").replace("http://", "https://")
        for x in links
        if "camara" in x
    ]
    senado = [
        x.replace("http://", "https://")
        for x in links
        if "senado.leg.br/norma/" in x and "/publicacao/" in x
    ]
    planalto = next((x for x in links if "planalto" in x), "-")
    alvos: list[tuple[str, str]] = []
    if camara:
        alvos.append(
            ("atualizada", re.sub(r"-norma-(p[el])\.html$", r"-normaatualizada-\1.html", camara[0]))
        )
        alvos += [("original", x) for x in camara if "publicacaooriginal" in x][:1]
    alvos += [("original_senado", x) for x in senado][:1]
    feitos = []
    for tipo, url in alvos:
        try:
            t = _texto_html(_decodifica(_get(url)))
        except OSError as e:
            feitos.append(f"{tipo}: erro {e}")
            continue
        (destino / f"{nome}__{tipo}.txt").write_text(
            _cabecalho(url, urn=urn, planalto=planalto, tipo=tipo) + t, encoding="utf-8"
        )
        feitos.append(f"{tipo}: {len(t)}")
    return feitos


def sijut(nome: str, id_ato: str, destino: Path) -> list[str]:
    d = json.loads(_get(f"{SIJUT}/{id_ato}/visao/multivigente"))

    def limpa(t: str | None) -> str:
        return html.unescape(re.sub(r"<[^>]+>", "", t or "")).strip()

    linhas = [
        f"ATO: {d['epigrafeCompleta']}",
        f"PUBLICAÇÃO: {d.get('dadosPublicacao')}",
        f"VIGÊNCIA INÍCIO: {d.get('dataVigenciaInicio')}  VIGENTE: {d.get('vigente')}",
        "HISTÓRICO:",
        *[
            f"  - {h['texto']} vigência {h.get('dataInicioVigencia')}"
            for h in d.get("historico") or []
        ],
        "EMENTA: " + " ".join(limpa(e["textoIntegra"]) for e in d.get("ementas") or []),
        "---- TEXTO (linhas [REVOGADO/ANTERIOR] = redação tachada) ----",
    ]
    segmentos = sorted(d.get("outrosSegmentos") or [], key=lambda s: s["ordemSegmentoAto"])
    anexos = []
    for s in segmentos:
        notas = s.get("ancorasDestino") or []
        nota = " ".join(limpa(a.get("texto")) for a in (s.get("ancorasOrigem") or []) + notas)
        if notas:
            nota += f" vig. {notas[0].get('dataInicioVigencia')}"
        t = limpa(s.get("textoIntegra"))
        if t:
            linhas.append(
                ("[REVOGADO/ANTERIOR] " if s.get("tachado") else "")
                + t
                + (f"  {{{nota}}}" if nota else "")
            )
        if s.get("arquivoBinario"):
            anexos.append(s["arquivoBinario"])
    for a in anexos:
        url = f"{SIJUT}/{id_ato}/anexo/{a['idArquivoBinario']}"
        b = _get(url)
        corpo = _pdf_texto(b) if b[:4] == b"%PDF" else _decodifica(b)
        linhas.append(f"\n==== ANEXO {a.get('nomeArquivoBinario')} ({url}) ====\n{corpo}")
    pagina = f"https://normas.receita.fazenda.gov.br/sijut2consulta/link.action?idAto={id_ato}"
    (destino / f"{nome}.txt").write_text(
        _cabecalho(f"{SIJUT}/{id_ato}/visao/multivigente", pagina=pagina) + "\n".join(linhas),
        encoding="utf-8",
    )
    return [f"{len(segmentos)} segmentos, {len(anexos)} anexos"]


def pdf(nome: str, url: str, destino: Path) -> list[str]:
    t = _pdf_texto(_get(url, 300))
    (destino / f"{nome}.txt").write_text(_cabecalho(url) + t, encoding="utf-8")
    return [f"{len(t)}"]


def bcb(nome: str, origem: str, destino: Path) -> list[str]:
    tipo, numero = origem.split("|")
    url = (
        "https://www.bcb.gov.br/api/conteudo/app/normativos/exibenormativo?p1="
        f"{urllib.parse.quote(tipo)}&p2={numero}"
    )
    c = json.loads(_get(url))["conteudo"][0]
    t = _texto_html(c.get("Texto", ""))
    (destino / f"{nome}.txt").write_text(
        _cabecalho(url, titulo=str(c.get("Titulo")), situacao=str(c.get("Situacao"))) + t,
        encoding="utf-8",
    )
    return [f"{len(t)}"]


def govbr(nome: str, raiz: str, destino: Path, limite: int = 150) -> list[str]:
    pasta = destino / nome
    pasta.mkdir(exist_ok=True)
    fila, vistos = [raiz], set()
    while fila and len(vistos) < limite:
        url = fila.pop(0).rstrip("/")
        if url in vistos:
            continue
        vistos.add(url)
        try:
            h = _decodifica(_get(url))
        except OSError:
            continue
        i, j = h.find('id="content"'), h.find('id="portal-footer')
        corpo = re.sub(r"(?is)<nav.*?</nav>|<header.*?</header>", "", h[i:j] if i > 0 else h)
        arquivo = url[len(raiz) :].strip("/").replace("/", "__") or "raiz"
        (pasta / f"{arquivo}.txt").write_text(
            _cabecalho(url) + _texto_html(corpo), encoding="utf-8"
        )
        for link in re.findall(r'href="([^"#?]+)"', h):
            link = html.unescape(link).rstrip("/")
            if link.startswith(raiz) and not re.search(r"\.(pdf|png|jpg)|/view$|@@|RSS|atom", link):
                fila.append(link)
    return [f"{len(vistos)} páginas"]


BAIXADORES: dict[str, Callable[[str, str, Path], list[str]]] = {
    "lei": lei,
    "sijut": sijut,
    "pdf": pdf,
    "bcb": bcb,
    "govbr": govbr,
}


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="Baixa a íntegra das normas da base fiscal (docs/normas/acervo.csv)."
    )
    p.add_argument("--destino", type=Path, default=DESTINO)
    p.add_argument("--so", nargs="*", help="nomes da coluna 'nome' de acervo.csv")
    a = p.parse_args(argv)
    a.destino.mkdir(parents=True, exist_ok=True)
    falhas = 0
    with LISTA.open(encoding="utf-8", newline="") as f:
        for linha in csv.DictReader(f):
            if a.so and linha["nome"] not in a.so:
                continue
            try:
                r = BAIXADORES[linha["tipo"]](linha["nome"], linha["origem"], a.destino)
            except (OSError, ValueError, KeyError) as e:
                r, falhas = [f"ERRO {e}"], falhas + 1
            print(linha["nome"], "|", "; ".join(r))
    return 1 if falhas else 0


if __name__ == "__main__":
    sys.exit(main())
