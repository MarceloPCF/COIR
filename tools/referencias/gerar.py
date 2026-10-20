"""Gerador das tabelas de referência públicas (F7b): respostas brutas das fontes -> CSV de data/.

Puro e determinístico: mesma entrada, mesmos bytes (linhas ordenadas, sem relógio; a data da coleta
entra como parâmetro). Só a biblioteca padrão (roda no GitHub Actions sem instalar o COIR).

Arquivos gerados (os curados à mão, como ``ativos.csv``, continuam valendo por cima deles):
- ``instrumentos_b3.csv``: ações, units, BDR, ETF e fundos da B3 com ISIN (histórico: primeira e
  última coleta em que o ticker apareceu; nunca apaga linha);
- ``ativos_b3.csv``: nome de pregão (nome da companhia + classe) -> ticker, como as notas imprimem;
- ``tipos_ativo_b3.csv``: UNIT, BDR, ETF e FII pela categoria da B3 (fundo só com "imobiliário" ou
  "Fiagro" no nome; FIP, FIDC e outros ficam de fora: sem palpite);
- ``opcoes_raiz_b3.csv``: raiz de 4 letras -> ativo-objeto (só raiz com um objeto só);
- ``eventos_b3.csv``: desdobramentos, grupamentos, bonificações, cisões e subscrições (nunca apaga);
- ``selic_mensal.csv``: série 4390 do BCB, só meses encerrados;
- ``pacote.json``: formato e data do pacote.
"""

from __future__ import annotations

import csv
import io
import json
import re
import unicodedata
from collections import defaultdict
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

FORMATO = 1  # esquema dos arquivos de data/; o app só aceita pacote do mesmo formato
CAIXA = "CASH"
CATEGORIAS = {"SHARES", "UNIT", "BDR", "FUNDS", "ETF EQUITIES", "ETF FOREIGN INDEX"}
TIPO_POR_CATEGORIA = {"UNIT": "UNIT", "BDR": "BDR", "ETF EQUITIES": "ETF",
                      "ETF FOREIGN INDEX": "ETF"}  # fmt: skip
OPCAO = "OPTION ON EQUITIES"
RE_RAIZ = re.compile(r"^[A-Z0-9]{4}")
# ação (final 3 a 8) e unit (11); recibos (9, 10), direitos (1, 2, 12...) e lotes especiais, fora
RE_ACAO = re.compile(r"^[A-Z0-9]{4}[3-8]$")
RE_UNIT = re.compile(r"^[A-Z0-9]{4}11$")


def _padrao(ticker: str, categoria: str = "") -> bool:
    if categoria == "UNIT":
        return bool(RE_UNIT.match(ticker))
    if categoria == "SHARES":
        return bool(RE_ACAO.match(ticker))
    return bool(RE_ACAO.match(ticker) or RE_UNIT.match(ticker))


COLUNAS_INSTRUMENTOS = ("ticker", "isin", "categoria", "especificacao", "emissor", "nome",
                        "visto_desde", "visto_ate")  # fmt: skip
COLUNAS_EVENTOS = ("emissor", "isin", "tipo", "fator", "data_com", "aprovado_em", "isin_emitido",
                   "preco", "periodo", "data_subscricao")  # fmt: skip


class GeracaoError(Exception):
    """Resposta de fonte fora do formato esperado: o pacote não é gerado (falha alto)."""


def _sem_acento(texto: str) -> str:
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode().upper()


def _br(texto: str) -> Decimal:
    """'1.234,50000000000' -> Decimal('1234.5')."""
    try:
        return Decimal(texto.strip().replace(".", "").replace(",", ".")).normalize()
    except InvalidOperation as erro:
        raise GeracaoError(f"número ilegível: {texto!r}") from erro


def _data_br(texto: str) -> str:
    """'08/08/2025' -> '2025-08-08' (vazio fica vazio)."""
    texto = texto.strip()
    if not texto:
        return ""
    try:
        d, m, a = texto.split("/")
        dia = date(int(a), int(m), int(d))
    except ValueError as erro:
        raise GeracaoError(f"data ilegível: {texto!r}") from erro
    return "" if dia.year in (1900, 9999) else dia.isoformat()  # datas fictícias da B3


def _periodo(texto: str) -> str:
    """'11/03/2026 a 27/04/2026'; o fictício '01/01/1900 a 01/01/1900' fica vazio."""
    return "" if "1900" in texto or "9999" in texto else " ".join(texto.split())


def _numero(d: Decimal) -> str:
    return format(d, "f")


# --------------------------------------------------------------------------------- leitura
@dataclass(frozen=True)
class Instrumento:
    ticker: str
    isin: str
    categoria: str
    especificacao: str  # "ON NM", "PN N1", "UNT N2", "CI", "DRN"
    emissor: str  # código de 4 letras da companhia (Asst)
    nome: str  # razão social


@dataclass(frozen=True)
class Opcao:
    ticker: str
    objeto: str


def ler_instrumentos(bruto: bytes) -> tuple[list[Instrumento], list[Opcao]]:
    """Cadastro de Instrumentos consolidado (CSV ';' em latin-1, 1ª linha de status)."""
    texto = bruto.decode("latin-1")
    linhas = texto.splitlines()
    if len(linhas) < 2 or not linhas[0].startswith("Status do Arquivo"):
        raise GeracaoError("cadastro de instrumentos: cabeçalho inesperado")
    leitor = csv.DictReader(io.StringIO("\n".join(linhas[1:])), delimiter=";")
    faltam = {"TckrSymb", "Asst", "SgmtNm", "SctyCtgyNm", "ISIN", "SpcfctnCd", "CrpnNm"} - set(
        leitor.fieldnames or ()
    )
    if faltam:
        raise GeracaoError(f"cadastro de instrumentos sem as colunas {sorted(faltam)}")
    instrumentos: list[Instrumento] = []
    opcoes: list[Opcao] = []
    for r in leitor:
        if r["SctyCtgyNm"] == OPCAO and r["Asst"]:
            opcoes.append(Opcao(r["TckrSymb"], r["Asst"]))
        elif r["SgmtNm"] == CAIXA and r["SctyCtgyNm"] in CATEGORIAS and r["ISIN"]:
            instrumentos.append(Instrumento(
                ticker=r["TckrSymb"].strip(), isin=r["ISIN"].strip(),
                categoria=r["SctyCtgyNm"], especificacao=" ".join(r["SpcfctnCd"].split()),
                emissor=r["Asst"].strip(), nome=" ".join(r["CrpnNm"].split()),
            ))  # fmt: skip
    if not instrumentos:
        raise GeracaoError("cadastro de instrumentos sem ações nem fundos")
    return instrumentos, opcoes


def nomes_de_pregao(empresas: Iterable[Mapping[str, object]]) -> dict[str, str]:
    """Lista de companhias (GetInitialCompanies) -> código do emissor -> nome de pregão."""
    saida: dict[str, str] = {}
    for e in empresas:
        codigo, nome = str(e.get("issuingCompany", "")).strip(), str(e.get("tradingName", ""))
        if codigo and nome.strip():
            saida[codigo] = " ".join(nome.split())
    return saida


@dataclass(frozen=True)
class EventoB3:
    emissor: str
    isin: str
    tipo: str  # como a B3 escreve: DESDOBRAMENTO, GRUPAMENTO, BONIFICACAO, CIS RED CAP, SUBSCRICAO
    fator: str  # % (desdobramento 200 = 1 vira 3; grupamento 0,1 = 10 viram 1); subscrição: %
    data_com: str
    aprovado_em: str
    isin_emitido: str
    preco: str = ""
    periodo: str = ""
    data_subscricao: str = ""

    def linha(self) -> dict[str, str]:
        return {c: getattr(self, c) for c in COLUNAS_EVENTOS}


def ler_eventos(emissor: str, resposta: Sequence[Mapping[str, object]]) -> list[EventoB3]:
    """GetListedSupplementCompany de uma companhia -> eventos em ações e subscrições."""
    saida: list[EventoB3] = []
    for bloco in resposta:
        for e in _lista(bloco, "stockDividends"):
            saida.append(EventoB3(
                emissor=emissor, isin=str(e["isinCode"]), tipo=str(e["label"]).strip(),
                fator=_numero(_br(str(e["factor"]))), data_com=_data_br(str(e["lastDatePrior"])),
                aprovado_em=_data_br(str(e.get("approvedOn", ""))),
                isin_emitido=str(e.get("assetIssued", "")),
            ))  # fmt: skip
        for e in _lista(bloco, "subscriptions"):
            saida.append(EventoB3(
                emissor=emissor, isin=str(e["isinCode"]), tipo=str(e["label"]).strip(),
                fator=_numero(_br(str(e.get("percentage", "0")))),
                data_com=_data_br(str(e["lastDatePrior"])),
                aprovado_em=_data_br(str(e.get("approvedOn", ""))),
                isin_emitido=str(e.get("assetIssued", "")),
                preco=_numero(_br(str(e.get("priceUnit", "0")))),
                periodo=_periodo(str(e.get("tradingPeriod", ""))),
                data_subscricao=_data_br(str(e.get("subscriptionDate", ""))),
            ))  # fmt: skip
    return saida


def _lista(bloco: Mapping[str, object], chave: str) -> list[Mapping[str, object]]:
    valor = bloco.get(chave) or []
    if not isinstance(valor, list):
        raise GeracaoError(f"eventos: {chave} não é lista")
    return [x for x in valor if isinstance(x, dict)]  # pyright: ignore[reportUnknownVariableType]


def ler_selic(resposta: Sequence[Mapping[str, str]], ate: date) -> dict[str, str]:
    """Série 4390 (% a.m.) -> mês AAAA-MM -> taxa, só meses encerrados antes de ``ate``."""
    saida: dict[str, str] = {}
    for r in resposta:
        _, m, a = r["data"].split("/")
        mes = f"{a}-{m}"
        if (int(a), int(m)) < (ate.year, ate.month):
            saida[mes] = _numero(Decimal(r["valor"]).quantize(Decimal("0.01")))
    return saida


# ---------------------------------------------------------------------------- tabelas
def ler_csv(caminho: Path) -> list[dict[str, str]]:
    if not caminho.is_file():
        return []
    with caminho.open(encoding="utf-8", newline="") as f:
        return [dict(r) for r in csv.DictReader(f)]


def texto_csv(colunas: Sequence[str], linhas: Iterable[Mapping[str, str]]) -> str:
    saida = io.StringIO()
    escritor = csv.DictWriter(saida, fieldnames=list(colunas), lineterminator="\n")
    escritor.writeheader()
    for linha in linhas:
        escritor.writerow({c: linha.get(c, "") for c in colunas})
    return saida.getvalue()


@dataclass
class Resumo:
    ativos_novos: list[str] = field(default_factory=list[str])
    trocas_candidatas: list[tuple[str, str, str]] = field(
        default_factory=list[tuple[str, str, str]]
    )  # (antigo, novo, visto até)
    eventos_novos: list[str] = field(default_factory=list[str])
    selic_nova: list[str] = field(default_factory=list[str])
    conflitos: list[str] = field(default_factory=list[str])
    sem_resposta: list[str] = field(default_factory=list[str])

    def markdown(self, data: date) -> str:
        partes = [f"## Referências públicas — coleta de {data:%d/%m/%Y}", ""]
        for titulo, itens in (
            ("Ativos novos", self.ativos_novos),
            (
                "Trocas de ticker candidatas (mesmo ISIN; revise e inclua em trocas_ticker.csv)",
                [f"{a} -> {n} (até {ate})" for a, n, ate in self.trocas_candidatas],
            ),
            ("Eventos corporativos novos", self.eventos_novos),
            ("Meses de Selic novos ou revistos", self.selic_nova),
            ("Ignorados por conflito (revise)", self.conflitos),
            (
                "Companhias sem resposta da B3 nesta coleta (eventos delas não atualizados)",
                self.sem_resposta,
            ),
        ):
            partes += [f"### {titulo} ({len(itens)})", *(f"- {i}" for i in itens[:200]), ""]
        return "\n".join(partes)

    @property
    def vazio(self) -> bool:
        return not (self.ativos_novos or self.trocas_candidatas or self.eventos_novos
                    or self.selic_nova)  # fmt: skip


def atualizar_instrumentos(
    anteriores: list[dict[str, str]], atuais: list[Instrumento], data: date, resumo: Resumo
) -> list[dict[str, str]]:
    """Junta a coleta ao histórico: atualiza ``visto_ate`` de quem apareceu, inclui os novos e
    nunca apaga (ticker que sumiu continua, com a última data vista)."""
    hoje = data.isoformat()
    por_ticker = {r["ticker"]: dict(r) for r in anteriores}
    vistos: set[str] = set()
    for i in sorted(atuais, key=lambda x: x.ticker):
        if i.ticker in vistos:
            continue
        vistos.add(i.ticker)
        linha = por_ticker.get(i.ticker)
        if linha is None:
            resumo.ativos_novos.append(f"{i.ticker} ({i.categoria}, {i.isin})")
            linha = {"ticker": i.ticker, "visto_desde": hoje}
            por_ticker[i.ticker] = linha
        linha.update(isin=i.isin, categoria=i.categoria, especificacao=i.especificacao,
                     emissor=i.emissor, nome=i.nome, visto_ate=hoje)  # fmt: skip
    # mesmo ISIN, ticker novo nesta coleta e o antigo sumiu: candidata a troca de ticker
    antigos_por_isin: dict[str, list[dict[str, str]]] = defaultdict(list)
    for r in anteriores:
        if r["ticker"] not in vistos:
            antigos_por_isin[r["isin"]].append(r)
    for i in atuais:
        if any(r["ticker"] == i.ticker for r in anteriores):
            continue
        for antigo in antigos_por_isin.get(i.isin, []):
            resumo.trocas_candidatas.append((antigo["ticker"], i.ticker, antigo["visto_ate"]))
    return [por_ticker[t] for t in sorted(por_ticker)]


def ativos(instrumentos: list[dict[str, str]], nomes: Mapping[str, str],
           resumo: Resumo) -> list[dict[str, str]]:  # fmt: skip
    """Nome de pregão + classe -> ticker (só ação e unit vigentes com nome de pregão conhecido).
    Nome que levaria a dois tickers fica de fora (conflito no resumo)."""
    candidatos: dict[str, set[str]] = defaultdict(set)
    ultima = max((r["visto_ate"] for r in instrumentos), default="")
    for r in instrumentos:
        if (r["categoria"] not in ("SHARES", "UNIT") or r["visto_ate"] != ultima
                or not _padrao(r["ticker"], r["categoria"])):  # fmt: skip
            continue
        nome, classe = nomes.get(r["emissor"]), (r["especificacao"].split() or [""])[0]
        if nome and classe:
            candidatos[f"{nome} {classe}"].add(r["ticker"])
    saida: list[dict[str, str]] = []
    for nome in sorted(candidatos):
        tickers = sorted(candidatos[nome])
        if len(tickers) > 1:
            resumo.conflitos.append(f"nome de pregão {nome!r}: {', '.join(tickers)}")
            continue
        saida.append({"nome_pregao": nome, "ticker": tickers[0]})
    return saida


def tipos(instrumentos: list[dict[str, str]]) -> list[dict[str, str]]:
    saida: list[dict[str, str]] = []
    for r in instrumentos:
        tipo = TIPO_POR_CATEGORIA.get(r["categoria"])
        if r["categoria"] == "FUNDS":
            nome = _sem_acento(r["nome"])
            imobiliario = "IMOB" in nome or "FIAGRO" in nome
            tipo = "FII" if imobiliario and " PART" not in nome else None
        if tipo is not None:
            saida.append(
                {"ticker": r["ticker"], "tipo": tipo, "observacao": "B3: " + r["categoria"]}
            )
    return saida


def opcoes_raiz(opcoes: Iterable[Opcao], resumo: Resumo) -> list[dict[str, str]]:
    objetos: dict[str, set[str]] = defaultdict(set)
    for o in opcoes:
        m = RE_RAIZ.match(o.ticker)
        if m and _padrao(o.objeto):
            objetos[m.group()].add(o.objeto)
    saida: list[dict[str, str]] = []
    for raiz in sorted(objetos):
        if len(objetos[raiz]) == 1:
            saida.append({"raiz": raiz, "ticker": next(iter(objetos[raiz]))})
        else:
            resumo.conflitos.append(f"raiz de opção {raiz}: {', '.join(sorted(objetos[raiz]))}")
    return saida


def juntar_eventos(anteriores: list[dict[str, str]], novos: Iterable[EventoB3],
                   resumo: Resumo) -> list[dict[str, str]]:  # fmt: skip
    """União com o histórico (a B3 tira da lista eventos de companhias deslistadas)."""

    def chave(r: Mapping[str, str]) -> tuple[str, ...]:
        return tuple(r[c] for c in ("isin", "tipo", "data_com", "isin_emitido", "fator"))

    todos = {chave(r): dict(r) for r in anteriores}
    for e in novos:
        linha = e.linha()
        if not e.data_com or not e.isin or Decimal(e.fator) <= 0:
            resumo.conflitos.append(f"evento sem data com, ISIN ou fator: {e.emissor} {e.tipo}")
            continue
        if chave(linha) not in todos:
            resumo.eventos_novos.append(f"{e.data_com} {e.emissor} {e.tipo} {e.fator}%")
            todos[chave(linha)] = linha
    return [todos[k] for k in sorted(todos, key=lambda k: (k[2], k[0], k[1], k[3], k[4]))]


def juntar_selic(anteriores: list[dict[str, str]], novos: Mapping[str, str],
                 resumo: Resumo) -> list[dict[str, str]]:  # fmt: skip
    """Meses novos (e revisões do BCB) a partir do primeiro mês já na tabela."""
    taxas = {r["mes"]: r["taxa"] for r in anteriores}
    inicio = min(taxas, default=min(novos, default=""))
    for mes, taxa in sorted(novos.items()):
        if mes >= inicio and taxas.get(mes) != taxa:
            resumo.selic_nova.append(f"{mes}: {taxa}%" + (f" (era {taxas[mes]})" if mes in taxas
                                                         else ""))  # fmt: skip
            taxas[mes] = taxa
    return [{"mes": m, "taxa": taxas[m]} for m in sorted(taxas)]


@dataclass(frozen=True)
class Coleta:
    """O que as fontes devolveram numa coleta (bruto, para o gerador ser testável sem rede)."""

    data: date
    instrumentos: bytes
    empresas: list[Mapping[str, object]]
    eventos: dict[str, list[Mapping[str, object]]]  # emissor -> resposta
    selic: list[Mapping[str, str]]
    sem_resposta: tuple[str, ...] = ()  # emissores que a B3 não respondeu nesta coleta

    def salvar(self, pasta: Path) -> None:
        """Respostas brutas em ``pasta`` (para reprocessar sem rede e depurar)."""
        pasta.mkdir(parents=True, exist_ok=True)
        (pasta / "instrumentos.csv").write_bytes(self.instrumentos)
        for nome, valor in (("empresas", self.empresas), ("eventos", self.eventos),
                            ("selic", self.selic), ("data", self.data.isoformat()),
                            ("sem_resposta", list(self.sem_resposta))):  # fmt: skip
            (pasta / f"{nome}.json").write_text(json.dumps(valor, ensure_ascii=False),
                                                encoding="utf-8")  # fmt: skip

    @classmethod
    def carregar(cls, pasta: Path) -> Coleta:
        def ler(nome: str) -> Any:
            return json.loads((pasta / f"{nome}.json").read_text(encoding="utf-8"))

        return cls(data=date.fromisoformat(ler("data")),
                   instrumentos=(pasta / "instrumentos.csv").read_bytes(),
                   empresas=ler("empresas"), eventos=ler("eventos"), selic=ler("selic"),
                   sem_resposta=tuple(ler("sem_resposta")))  # fmt: skip


RESPOSTA_RUIM = (KeyError, IndexError, TypeError, ValueError, AttributeError, InvalidOperation)


def gerar(pasta: Path, coleta: Coleta) -> tuple[dict[str, str], Resumo]:
    """Novos conteúdos dos arquivos gerados em ``pasta`` (data/) e o resumo das mudanças. Resposta
    fora do formato (campo faltando, número ou data ilegível) falha alto como GeracaoError."""
    try:
        return _gerar(pasta, coleta)
    except GeracaoError:
        raise
    except RESPOSTA_RUIM as erro:
        raise GeracaoError(f"resposta de fonte fora do formato: {erro!r}") from erro


def _gerar(pasta: Path, coleta: Coleta) -> tuple[dict[str, str], Resumo]:
    resumo = Resumo(sem_resposta=list(coleta.sem_resposta))
    atuais, opcoes = ler_instrumentos(coleta.instrumentos)
    historico = atualizar_instrumentos(ler_csv(pasta / "instrumentos_b3.csv"), atuais,
                                       coleta.data, resumo)  # fmt: skip
    nomes = nomes_de_pregao(coleta.empresas)
    eventos = [e for emissor, resp in sorted(coleta.eventos.items())
               for e in ler_eventos(emissor, resp)]  # fmt: skip
    arquivos = {
        "instrumentos_b3.csv": texto_csv(COLUNAS_INSTRUMENTOS, historico),
        "ativos_b3.csv": texto_csv(("nome_pregao", "ticker"), ativos(historico, nomes, resumo)),
        "tipos_ativo_b3.csv": texto_csv(("ticker", "tipo", "observacao"), tipos(historico)),
        "opcoes_raiz_b3.csv": texto_csv(("raiz", "ticker"), opcoes_raiz(opcoes, resumo)),
        "eventos_b3.csv": texto_csv(COLUNAS_EVENTOS, juntar_eventos(
            ler_csv(pasta / "eventos_b3.csv"), eventos, resumo)),
        "selic_mensal.csv": texto_csv(("mes", "taxa"), juntar_selic(
            ler_csv(pasta / "selic_mensal.csv"), ler_selic(coleta.selic, coleta.data), resumo)),
    }  # fmt: skip
    arquivos["pacote.json"] = json.dumps({"formato": FORMATO, "data": coleta.data.isoformat()},
                                         indent=1) + "\n"  # fmt: skip
    return arquivos, resumo


def gravar(pasta: Path, arquivos: Mapping[str, str]) -> list[str]:
    """Grava os arquivos que mudaram (devolve os nomes). ``pacote.json`` só muda com o resto."""
    mudaram = [
        n
        for n, t in arquivos.items()
        if n != "pacote.json"
        and (not (pasta / n).is_file() or (pasta / n).read_text(encoding="utf-8") != t)
    ]
    for nome in [*mudaram, *(["pacote.json"] if mudaram else [])]:
        (pasta / nome).write_text(arquivos[nome], encoding="utf-8", newline="")
    return mudaram
