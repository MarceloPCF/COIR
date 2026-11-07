"""Gerador das tabelas de referência públicas (F7b): respostas brutas das fontes -> CSV de data/.

Puro e determinístico: mesma entrada, mesmos bytes (linhas ordenadas, sem relógio; a data da coleta
entra como parâmetro). Só a biblioteca padrão (roda no GitHub Actions sem instalar o COIR).

Arquivos gerados (os curados à mão, como ``ativos.csv``, continuam valendo por cima deles):
- ``instrumentos_b3.csv``: ações, units, BDR, ETF e fundos da B3 com ISIN (histórico: primeira e
  última coleta em que o ticker apareceu; nunca apaga linha);
- ``ativos_b3.csv``: nome de pregão (nome da companhia + classe) -> ticker, como as notas imprimem;
- ``nomes_pregao_b3.csv``: nome de pregão -> ticker com o período, do COTAHIST (como as notas
  imprimem, de 2008 em diante) e da lista de ETF de renda fixa (fora do COTAHIST); nunca apaga;
- ``tipos_ativo_b3.csv``: UNIT, BDR, ETF e FII pela categoria da B3 (fundo só com "imobiliário" ou
  "Fiagro" no nome; FIP, FIDC e outros ficam de fora: sem palpite);
- ``opcoes_raiz_b3.csv``: raiz de 4 letras -> ativo-objeto (só raiz com um objeto só);
- ``eventos_b3.csv``: desdobramentos, grupamentos, bonificações, cisões e subscrições (nunca apaga),
  com a origem de cada linha (``b3``, ``cvm_fre``, ``cvm_ipe``; ``b3+cvm_fre`` quando as duas
  concordam); divergência entre fontes não entra e vai para o resumo do PR (``historico.py``);
- ``selic_mensal.csv``: série 4390 do BCB, só meses encerrados;
- ``proventos_b3.csv``: proventos em dinheiro (dividendo, JCP, rendimento) com data de aprovação,
  data com, pagamento e valor por ação (o crédito do JCP; a B3 só lista os últimos meses, nunca
  apaga);
- ``proventos_historico_b3.csv``: JCP da aba "Proventos em dinheiro" de cada companhia desde
  ``HISTORICO_DESDE``, com classe, data de aprovação e data com (sem pagamento nem ISIN: a B3 não
  publica); cobre os anos que ``proventos_b3.csv`` não alcança; nunca apaga;
- ``emissores_b3.csv``: CNPJ e razão social de cada emissor (companhias e fundos FII, Fiagro e ETF),
  para o Bens e Direitos do IRPF (histórico: o CNPJ de um fundo pode mudar; nunca apaga linha);
- ``pacote.json``: formato e data do pacote.
"""

from __future__ import annotations

import csv
import io
import json
import re
import unicodedata
from collections import defaultdict
from collections.abc import Collection, Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
from fractions import Fraction
from pathlib import Path
from typing import Any

FORMATO = 2  # esquema dos arquivos de data/; o app só aceita pacote do mesmo formato
CAIXA = "CASH"
CATEGORIAS = {"SHARES", "UNIT", "BDR", "FUNDS", "ETF EQUITIES", "ETF FOREIGN INDEX"}
# ETF de renda fixa: a B3 os publica fora do segmento à vista, sem especificação nem razão social
CATEGORIA_RF = "FIXED INCOME TRADABLE INSTRUMENT T1"
SEGMENTO_RF = "FORWARD"
# cancelado, sem índice nem carteira (parecer do especialista de 07/11/2026): sem tipo, sem palpite
FORA_DO_ETF_RF = {"GPRE11"}
TIPO_POR_CATEGORIA = {"UNIT": "UNIT", "BDR": "BDR", "ETF EQUITIES": "ETF",
                      "ETF FOREIGN INDEX": "ETF", CATEGORIA_RF: "ETF_RF"}  # fmt: skip
OPCAO = "OPTION ON EQUITIES"
RE_RAIZ = re.compile(r"^[A-Z0-9]{4}")
# ação (final 3 a 8) e unit (11); recibos (9, 10), direitos (1, 2, 12...) e lotes especiais, fora
RE_ACAO = re.compile(r"^[A-Z0-9]{4}[3-8]$")
RE_UNIT = re.compile(r"^[A-Z0-9]{4}11$")
# ação com final fora do padrão (RAIZ55 "PNA P N2": lote de cotação): a classe da B3 diz que é ação
RE_ACAO_ESPECIAL = re.compile(r"^[A-Z0-9]{4}\d{2,3}$")
CLASSES_ACAO = {"ON", "PN", "PNA", "PNB", "PNC", "PND", "PNE", "PNF", "PNG"}
# ativo-objeto de opção com final especial: só as ações "P" (RAIZ55); AUAU99 e semelhantes, não
RE_OBJETO_ESPECIAL = re.compile(r"^[A-Z0-9]{4}5[3-6]$")


def _padrao(ticker: str, categoria: str = "") -> bool:
    if categoria == "UNIT":
        return bool(RE_UNIT.match(ticker))
    if categoria == "SHARES":
        return bool(RE_ACAO.match(ticker))
    return bool(RE_ACAO.match(ticker) or RE_UNIT.match(ticker))


COLUNAS_NOMES = ("nome_pregao", "ticker", "desde", "ate", "tipo")
COLUNAS_TROCAS = ("ticker_impresso", "ticker", "ate", "isin_antigo", "isin_novo", "fonte")
COLUNAS_INSTRUMENTOS = ("ticker", "isin", "categoria", "especificacao", "emissor", "nome",
                        "visto_desde", "visto_ate")  # fmt: skip
COLUNAS_EMISSORES = ("emissor", "cnpj", "razao_social", "tipo", "visto_desde", "visto_ate")
COLUNAS_PROVENTOS = ("emissor", "isin", "tipo", "valor", "aprovado_em", "data_com", "pagamento",
                     "referente")  # fmt: skip
COLUNAS_HISTORICO = ("emissor", "classe", "tipo", "valor", "aprovado_em", "data_com")
JCP = "JRS CAP PROPRIO"
HISTORICO_DESDE = "2016-01-01"  # data com mais antiga guardada (dez anos antes da F8)
COLUNAS_EVENTOS = ("emissor", "isin", "tipo", "fator", "data_com", "aprovado_em", "isin_emitido",
                   "preco", "periodo", "data_subscricao", "origem", "referencia",
                   "custo_atribuido")  # fmt: skip
TIPOS_COM_FATOR = ("DESDOBRAMENTO", "GRUPAMENTO", "BONIFICACAO")  # comparáveis entre fontes
MESMO_EVENTO = timedelta(days=30)  # mesma espécie, tipo e fator, data com tão perto: mesmo evento


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
        elif r["ISIN"] and (
            (r["SgmtNm"] == CAIXA and r["SctyCtgyNm"] in CATEGORIAS)
            or (r["SgmtNm"] == SEGMENTO_RF and r["SctyCtgyNm"] == CATEGORIA_RF)
        ):
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


def cnpj_valido(cnpj: str) -> bool:
    """14 dígitos com os dois dígitos verificadores certos (módulo 11). Cópia de
    coir.domain.cpf.cnpj_valido: o pipeline roda sem instalar o COIR (só a biblioteca padrão)."""
    if len(cnpj) != 14 or not cnpj.isdigit() or len(set(cnpj)) == 1:
        return False
    for n in (12, 13):
        pesos = list(range(n - 7, 1, -1)) + list(range(9, 1, -1))
        resto = sum(int(c) * p for c, p in zip(cnpj[:n], pesos, strict=True)) % 11
        if int(cnpj[n]) != (0 if resto < 2 else 11 - resto):
            return False
    return True


@dataclass(frozen=True)
class Emissor:
    emissor: str  # código de 4 letras (Asst no cadastro de instrumentos; sigla do fundo)
    cnpj: str
    razao_social: str
    tipo: str  # COMPANHIA, FII, FIAGRO, ETF, ETF-RF, ETF-FII


def ler_emissores(empresas: Iterable[Mapping[str, object]], fundos: Iterable[Mapping[str, object]],
                  resumo: Resumo) -> list[Emissor]:  # fmt: skip
    """Companhias (GetInitialCompanies) e fundos listados (GetDetailFund) -> CNPJ por emissor.
    Sem CNPJ (BDR: "0") fica de fora; CNPJ com dígito errado vai para os conflitos."""
    saida: list[Emissor] = []
    for origem, codigo, nome, tipo in (
        *((e, "issuingCompany", "companyName", None) for e in empresas),
        *((f, "acronym", "fundName", "typeName") for f in fundos),
    ):
        emissor = str(origem.get(codigo) or "").strip()
        cnpj = re.sub(r"\D", "", str(origem.get("cnpj") or ""))
        if not emissor or cnpj.strip("0") == "":
            continue
        cnpj = cnpj.zfill(14)  # a B3 tira os zeros à esquerda
        if not cnpj_valido(cnpj):
            resumo.conflitos.append(f"CNPJ inválido para {emissor}: {cnpj}")
            continue
        razao = " ".join(str(origem.get(nome) or "").split())
        classe = str(origem.get(tipo) or "") if tipo else "COMPANHIA"
        saida.append(Emissor(emissor=emissor, cnpj=cnpj, razao_social=razao, tipo=classe))
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
    origem: str = "b3"  # b3 (consulta de empresas), cvm_fre (FRE + COTAHIST), cvm_ipe (aviso)
    referencia: str = ""  # documento da CVM (FRE ou protocolo do aviso aos acionistas)
    custo_atribuido: str = ""  # R$ por ação bonificada, quando o aviso informa

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


def ler_proventos(emissor: str, resposta: Sequence[Mapping[str, object]]) -> list[dict[str, str]]:
    """Proventos em dinheiro (cashDividends) de uma companhia: dividendo, JCP, rendimento, com a
    data de aprovação (o crédito do JCP), a data com e o pagamento; valor por ação bruto."""
    saida: list[dict[str, str]] = []
    for bloco in resposta:
        for e in _lista(bloco, "cashDividends"):
            saida.append({
                "emissor": emissor, "isin": str(e["isinCode"]), "tipo": str(e["label"]).strip(),
                "valor": _numero(_br(str(e["rate"]))),
                "aprovado_em": _data_br(str(e.get("approvedOn", ""))),
                "data_com": _data_br(str(e.get("lastDatePrior", ""))),
                "pagamento": _data_br(str(e.get("paymentDate", ""))),
                "referente": " ".join(str(e.get("relatedTo") or "").split()),
            })  # fmt: skip
    return saida


def juntar_proventos(anteriores: list[dict[str, str]], novos: Iterable[dict[str, str]],
                     resumo: Resumo) -> list[dict[str, str]]:  # fmt: skip
    """União com o histórico: a B3 só lista os proventos pagos nos últimos meses."""

    def chave(r: Mapping[str, str]) -> tuple[str, ...]:
        return tuple(r[c] for c in COLUNAS_PROVENTOS[1:7])

    todos = {chave(r): dict(r) for r in anteriores}
    for r in novos:
        if not r["isin"] or not r["aprovado_em"]:
            continue  # sem data de aprovação não serve para o crédito do JCP
        if chave(r) not in todos:
            resumo.proventos_novos.append(f"{r['aprovado_em']} {r['emissor']} {r['tipo']}")
            todos[chave(r)] = r
    return [todos[k] for k in sorted(todos, key=lambda k: (k[3], k[0], k[1], k[5], k[2]))]


def ler_historico(emissor: str, resultados: Sequence[Mapping[str, object]]) -> list[dict[str, str]]:
    """Aba "Proventos em dinheiro" (GetListedCashDividends) de uma companhia -> JCP com data de
    aprovação e data com desde ``HISTORICO_DESDE``. Linha sem aprovação ou com valor por lote (a
    B3 cotava por mil ações nos anos 1990) fica de fora."""
    saida: list[dict[str, str]] = []
    for e in resultados:
        if str(e["corporateAction"]).strip() != JCP or str(e.get("quotedPerShares", "1")) != "1":
            continue
        aprovado = _data_br(str(e.get("dateApproval") or ""))
        data_com = _data_br(str(e["lastDatePriorEx"]))
        if aprovado and data_com >= HISTORICO_DESDE:
            saida.append({"emissor": emissor, "classe": str(e["typeStock"]).strip(), "tipo": JCP,
                          "valor": _numero(_br(str(e["valueCash"]))), "aprovado_em": aprovado,
                          "data_com": data_com})  # fmt: skip
    return saida


def juntar_historico(anteriores: list[dict[str, str]], novos: Iterable[dict[str, str]],
                     resumo: Resumo) -> list[dict[str, str]]:  # fmt: skip
    """União com o que já foi coletado (nunca apaga), ordenada por data com."""

    def chave(r: Mapping[str, str]) -> tuple[str, ...]:
        return tuple(r[c] for c in COLUNAS_HISTORICO)

    todos = {chave(r): dict(r) for r in anteriores}
    for r in novos:
        if chave(r) not in todos:
            resumo.historico_novos.append(f"{r['data_com']} {r['emissor']} {r['classe']}")
            todos[chave(r)] = r
    return [todos[k] for k in sorted(todos, key=lambda k: (k[5], k[0], k[1], k[4], k[3]))]


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
    emissores_novos: list[str] = field(default_factory=list[str])
    cnpj_mudou: list[str] = field(default_factory=list[str])
    proventos_novos: list[str] = field(default_factory=list[str])
    historico_novos: list[str] = field(default_factory=list[str])
    divergencias: list[str] = field(default_factory=list[str])
    revisao: list[str] = field(default_factory=list[str])
    conflitos: list[str] = field(default_factory=list[str])
    sem_resposta: list[str] = field(default_factory=list[str])
    saidas: list[str] = field(default_factory=list[str])  # sumiram do cadastro nesta coleta
    nomes_novos: list[str] = field(default_factory=list[str])

    def markdown(self, data: date) -> str:
        partes = [f"## Referências públicas — coleta de {data:%d/%m/%Y}", ""]
        for titulo, itens in (
            (
                "Ativos que saíram do cadastro: toda linha pede revisão (troca de código, "
                "incorporação, cancelamento de registro?); o que tiver evento vai para "
                "trocas_ticker.csv ou eventos_cvm_ipe.csv",
                self.saidas,
            ),
            ("Ativos novos", self.ativos_novos),
            (
                "Trocas de ticker (mesmo ISIN; já incluídas em trocas_ticker.csv, confira)",
                [f"{a} -> {n} (até {ate})" for a, n, ate in self.trocas_candidatas],
            ),
            ("Eventos corporativos novos", self.eventos_novos),
            ("Nomes de pregão novos", self.nomes_novos),
            ("Meses de Selic novos ou revistos", self.selic_nova),
            ("Emissores novos (CNPJ)", self.emissores_novos),
            ("CNPJ novo para emissor já conhecido (revise o Bens e Direitos)", self.cnpj_mudou),
            (
                "Proventos em dinheiro novos (data de aprovação, emissor, tipo)",
                self.proventos_novos,
            ),
            ("JCP novos no histórico (data com, emissor, classe)", self.historico_novos),
            (
                "Eventos com fontes divergentes (revise; nada foi alterado no pacote)",
                self.divergencias,
            ),
            ("Eventos da CVM sem data com ou fator conferível (revise)", self.revisao),
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
                    or self.selic_nova or self.emissores_novos or self.cnpj_mudou
                    or self.proventos_novos or self.historico_novos
                    or self.divergencias or self.saidas or self.nomes_novos)  # fmt: skip


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
    _saidas(anteriores, atuais, vistos, resumo)
    return [por_ticker[t] for t in sorted(por_ticker)]


def _saidas(anteriores: list[dict[str, str]], atuais: list[Instrumento], vistos: set[str],
            resumo: Resumo) -> None:  # fmt: skip
    """Quem estava no cadastro da coleta anterior e sumiu desta, com os códigos novos do mesmo
    emissor: nada sai sem alguém olhar (PETZ3 virou AUAU3, de outro emissor, sem sugestão)."""
    ultima = max((r["visto_ate"] for r in anteriores), default="")
    novos: dict[str, list[str]] = defaultdict(list)
    conhecidos = {r["ticker"] for r in anteriores}
    for i in atuais:
        if i.ticker not in conhecidos:
            novos[i.emissor].append(i.ticker)
    for r in anteriores:
        if r["visto_ate"] != ultima or r["ticker"] in vistos:
            continue
        sucessores = novos.get(r["emissor"], [])
        sugestao = (f"; novos do emissor: {', '.join(sucessores)}" if sucessores
                    else "; sem sucessor")  # fmt: skip
        resumo.saidas.append(f"{r['ticker']} ({r['especificacao']}, {r['nome']}){sugestao}")


def juntar_trocas(anteriores: list[dict[str, str]], candidatas: Iterable[tuple[str, str, str]],
                  instrumentos: list[dict[str, str]], nomes: Iterable[Mapping[str, str]],
                  ) -> list[dict[str, str]]:  # fmt: skip
    """Mesmo ISIN, código novo e o antigo fora do cadastro: é só troca de código, entra na tabela
    (o PR da coleta mostra a linha). ``ate`` é o último negócio no COTAHIST, se houver."""
    saida = [dict(r) for r in anteriores]
    ja = {r["ticker_impresso"] for r in saida}
    isin = {r["ticker"]: r["isin"] for r in instrumentos}
    ultimo: dict[str, str] = {}
    for n in nomes:
        ultimo[n["ticker"]] = max(n["ate"], ultimo.get(n["ticker"], ""))
    for antigo, novo, visto in candidatas:
        if antigo in ja:
            continue
        ja.add(antigo)
        ate = ultimo.get(antigo, visto)
        saida.append({"ticker_impresso": antigo, "ticker": novo, "ate": ate,
                      "isin_antigo": isin.get(antigo, ""), "isin_novo": isin.get(novo, ""),
                      "fonte": f"B3: mesmo ISIN no cadastro de instrumentos; último negócio de "
                               f"{antigo} em {ate}"})  # fmt: skip
    return saida


def juntar_nomes(anteriores: Iterable[Mapping[str, str]], novos: Iterable[Mapping[str, str]],
                 resumo: Resumo) -> list[dict[str, str]]:  # fmt: skip
    """União por (nome, ticker) com o período estendido; nunca apaga."""
    todos = {(r["nome_pregao"], r["ticker"]): dict(r) for r in anteriores}
    for r in novos:
        chave = (r["nome_pregao"], r["ticker"])
        linha = todos.get(chave)
        if linha is None:
            todos[chave] = dict(r)
            resumo.nomes_novos.append(f"{r['nome_pregao']} -> {r['ticker']}")
        else:
            linha["desde"] = min(linha["desde"], r["desde"])
            linha["ate"] = max(linha["ate"], r["ate"])
            linha["tipo"] = linha.get("tipo") or r.get("tipo", "")
    return [todos[k] for k in sorted(todos)]


def nomes_renda_fixa(fundos: Iterable[Mapping[str, object]], data: date) -> list[dict[str, str]]:
    """ETF de renda fixa não aparece no COTAHIST: nome de pregão e código vêm da lista de fundos da
    B3. A nota imprime "INVESTO LFTB F11" ("F11" é sufixo da especificação) ou "... CI"."""
    dia = data.isoformat()
    saida: list[dict[str, str]] = []
    for f in fundos:
        nome, codigo = " ".join(str(f.get("tradingName") or "").split()), f.get("tradingCode")
        if f.get("typeName") == "ETF-RF" and nome and codigo:
            for texto in (nome, f"{nome} CI"):
                saida.append({"nome_pregao": texto, "ticker": str(codigo), "desde": dia,
                              "ate": dia})  # fmt: skip
    return saida


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


def tipos(instrumentos: list[dict[str, str]],
          nomes: Iterable[Mapping[str, str]] = ()) -> list[dict[str, str]]:  # fmt: skip
    """Tipo pela categoria do cadastro; o ticker que já saiu do cadastro, pelo que o COTAHIST dizia
    dele (``nomes``: FII, Fiagro, unit e BDR). Fundo fora disso fica sem tipo: nada de palpite."""
    saida: list[dict[str, str]] = []
    vistos = {r["ticker"] for r in instrumentos}
    for r in instrumentos:
        tipo = TIPO_POR_CATEGORIA.get(r["categoria"])
        if r["ticker"] in FORA_DO_ETF_RF:
            continue
        if (r["categoria"] == "SHARES" and RE_ACAO_ESPECIAL.match(r["ticker"])
                and (r["especificacao"].split() or [""])[0] in CLASSES_ACAO):  # fmt: skip
            tipo = "ACAO"
        if r["categoria"] == "FUNDS":
            nome = _sem_acento(r["nome"])
            imobiliario = "IMOB" in nome or "FIAGRO" in nome
            tipo = "FII" if imobiliario and " PART" not in nome else None
        if tipo is not None:
            saida.append(
                {"ticker": r["ticker"], "tipo": tipo, "observacao": "B3: " + r["categoria"]}
            )
    antigos: dict[str, str] = {}
    for n in nomes:
        if n.get("tipo") and n["ticker"] not in vistos:
            antigos[n["ticker"]] = n["tipo"]
    saida += [{"ticker": t, "tipo": antigos[t], "observacao": "B3: COTAHIST"}
              for t in sorted(antigos)]  # fmt: skip
    return saida


def opcoes_raiz(opcoes: Iterable[Opcao], resumo: Resumo) -> list[dict[str, str]]:
    objetos: dict[str, set[str]] = defaultdict(set)
    for o in opcoes:
        m = RE_RAIZ.match(o.ticker)
        if m and (_padrao(o.objeto) or RE_OBJETO_ESPECIAL.match(o.objeto)):
            objetos[m.group()].add(o.objeto)
    saida: list[dict[str, str]] = []
    for raiz in sorted(objetos):
        if len(objetos[raiz]) == 1:
            saida.append({"raiz": raiz, "ticker": next(iter(objetos[raiz]))})
        else:
            resumo.conflitos.append(f"raiz de opção {raiz}: {', '.join(sorted(objetos[raiz]))}")
    return saida


def multiplicador(tipo: str, fator: str) -> Fraction:
    """Fator como a B3 escreve -> multiplicador da quantidade (grupamento 0,1 -> 1/10; os demais
    em %: desdobramento 200 -> 3)."""
    f = Fraction(Decimal(fator))
    return f if tipo == "GRUPAMENTO" else 1 + f / 100


def _equivalentes(a: Mapping[str, str], b: Mapping[str, str]) -> bool:
    """Mesmo fator a menos do arredondamento da dízima (33,333333333 x 33,333333334)."""
    fa, fb = multiplicador(a["tipo"], a["fator"]), multiplicador(b["tipo"], b["fator"])
    return abs(fa / fb - 1) < Fraction(1, 10**6)


def _familia(tipo: str) -> str:
    """Bonificação e desdobramento mexem na quantidade do mesmo jeito (uma fonte chama de um o que
    a outra chama de outro): no mesmo ISIN e data com, os dois juntos dobrariam a posição."""
    return "AUMENTO" if tipo in ("DESDOBRAMENTO", "BONIFICACAO") else tipo


def _somar_origem(linha: dict[str, str], nova: Mapping[str, str]) -> None:
    origens = set(linha["origem"].split("+")) | set(nova["origem"].split("+"))
    linha["origem"] = "+".join(sorted(origens))
    # por documento: rodar o histórico de novo não repete a referência que a linha já tem
    refs = [r for texto in (linha.get("referencia", ""), nova["referencia"])
            for r in texto.split("; ") if r]  # fmt: skip
    linha["referencia"] = "; ".join(dict.fromkeys(refs))
    linha["custo_atribuido"] = linha.get("custo_atribuido") or nova["custo_atribuido"]


def _texto_evento(r: Mapping[str, str]) -> str:
    return f"{r['data_com']} {r['emissor']} {r['isin']} {r['tipo']} {r['fator']} ({r['origem']})"


def juntar_eventos(anteriores: list[dict[str, str]], novos: Iterable[EventoB3], resumo: Resumo,
                   excluidos: Collection[tuple[str, ...]] = ()
                   ) -> list[dict[str, str]]:  # fmt: skip
    """União com o histórico (a B3 tira da lista eventos de companhias deslistadas e só guarda o
    último de cada tipo por ISIN). Nunca apaga sem registro. O mesmo evento vindo de outra fonte
    soma a origem; se a outra fonte traz outro fator na mesma data com, ou o mesmo fator em data
    com próxima, é divergência: fica como estava e vai para o resumo (revisão humana, sem
    palpite). O que está em ``excluidos`` (``eventos_excluidos.csv``, com o motivo) sai e não
    volta, mesmo que a B3 o devolva de novo."""

    def chave(r: Mapping[str, str]) -> tuple[str, ...]:
        return tuple(r[c] for c in COLUNAS_CHAVE)

    todos = {chave(r): {**r, "origem": r.get("origem") or "b3"} for r in anteriores
             if chave(r) not in excluidos}  # fmt: skip
    for e in novos:
        linha = e.linha()
        if chave(linha) in excluidos:
            continue
        if not e.data_com or not e.isin or Decimal(e.fator) <= 0:
            resumo.conflitos.append(f"evento sem data com, ISIN ou fator: {e.emissor} {e.tipo}")
            continue
        if chave(linha) in todos:
            _somar_origem(todos[chave(linha)], linha)
            continue
        if e.tipo in TIPOS_COM_FATOR:
            dia = date.fromisoformat(e.data_com)
            alvo = (e.isin, _familia(e.tipo), e.isin_emitido)
            vizinhos = [r for r in todos.values()
                        if (r["isin"], _familia(r["tipo"]), r["isin_emitido"]) == alvo
                        and set(r["origem"].split("+")) - {e.origem}
                        and abs(date.fromisoformat(r["data_com"]) - dia) <= MESMO_EVENTO
                        ]  # fmt: skip
            iguais = [r for r in vizinhos if r["tipo"] == e.tipo
                      and r["data_com"] == e.data_com and _equivalentes(r, linha)]  # fmt: skip
            if iguais:
                _somar_origem(iguais[0], linha)
                continue
            if any(r["data_com"] == e.data_com or _equivalentes(r, linha) for r in vizinhos):
                outras = ", ".join(_texto_evento(r) for r in vizinhos)
                resumo.divergencias.append(f"{_texto_evento(linha)} x {outras}")
                continue
        resumo.eventos_novos.append(f"{e.data_com} {e.emissor} {e.tipo} {e.fator}% ({e.origem})")
        todos[chave(linha)] = linha
    return [todos[k] for k in sorted(todos, key=lambda k: (k[2], k[0], k[1], k[3], k[4]))]


def juntar_emissores(anteriores: list[dict[str, str]], atuais: Iterable[Emissor], data: date,
                     resumo: Resumo) -> list[dict[str, str]]:  # fmt: skip
    """Histórico por (emissor, CNPJ): atualiza ``visto_ate`` e nunca apaga (o CNPJ de um fundo
    pode mudar; o Bens e Direitos de anos passados usa o que valia na época)."""
    hoje = data.isoformat()
    todos = {(r["emissor"], r["cnpj"]): dict(r) for r in anteriores}
    conhecidos = {r["emissor"] for r in anteriores}
    vistos: set[tuple[str, str]] = set()
    for e in sorted(atuais, key=lambda x: (x.emissor, x.cnpj, x.tipo)):
        chave = (e.emissor, e.cnpj)
        if chave in vistos:
            continue
        vistos.add(chave)
        linha = todos.get(chave)
        if linha is None:
            if e.emissor in conhecidos:
                resumo.cnpj_mudou.append(f"{e.emissor}: {e.cnpj} ({e.razao_social})")
            else:
                resumo.emissores_novos.append(f"{e.emissor} {e.cnpj} {e.razao_social}")
            linha = {"emissor": e.emissor, "cnpj": e.cnpj, "visto_desde": hoje}
            todos[chave] = linha
        linha.update(razao_social=e.razao_social, tipo=e.tipo, visto_ate=hoje)
    return [todos[k] for k in sorted(todos)]


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
    fundos: list[Mapping[str, object]] = field(default_factory=list[Mapping[str, object]])
    historico: dict[str, list[Mapping[str, object]]] = field(
        default_factory=dict[str, list[Mapping[str, object]]]
    )  # emissor -> linhas da aba "Proventos em dinheiro" (todas as páginas)
    nomes: Sequence[Mapping[str, str]] = ()  # do COTAHIST

    def salvar(self, pasta: Path) -> None:
        """Respostas brutas em ``pasta`` (para reprocessar sem rede e depurar)."""
        pasta.mkdir(parents=True, exist_ok=True)
        (pasta / "instrumentos.csv").write_bytes(self.instrumentos)
        for nome, valor in (("empresas", self.empresas), ("eventos", self.eventos),
                            ("selic", self.selic), ("data", self.data.isoformat()),
                            ("fundos", self.fundos), ("historico", self.historico),
                            ("sem_resposta", list(self.sem_resposta)),
                            ("nomes", list(self.nomes))):  # fmt: skip
            (pasta / f"{nome}.json").write_text(json.dumps(valor, ensure_ascii=False),
                                                encoding="utf-8")  # fmt: skip

    @classmethod
    def carregar(cls, pasta: Path) -> Coleta:
        def ler(nome: str) -> Any:
            return json.loads((pasta / f"{nome}.json").read_text(encoding="utf-8"))

        fundos = ler("fundos") if (pasta / "fundos.json").is_file() else []  # coleta antiga
        historico = ler("historico") if (pasta / "historico.json").is_file() else {}
        nomes = ler("nomes") if (pasta / "nomes.json").is_file() else []
        return cls(data=date.fromisoformat(ler("data")),
                   instrumentos=(pasta / "instrumentos.csv").read_bytes(),
                   empresas=ler("empresas"), eventos=ler("eventos"), selic=ler("selic"),
                   sem_resposta=tuple(ler("sem_resposta")), fundos=fundos,
                   historico=historico, nomes=nomes)  # fmt: skip


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
               for e in ler_eventos(emissor, resp)] + eventos_curados()  # fmt: skip
    excluidos = eventos_excluidos()
    arquivos = {
        "instrumentos_b3.csv": texto_csv(COLUNAS_INSTRUMENTOS, historico),
        "ativos_b3.csv": texto_csv(("nome_pregao", "ticker"), ativos(historico, nomes, resumo)),
        "opcoes_raiz_b3.csv": texto_csv(("raiz", "ticker"), opcoes_raiz(opcoes, resumo)),
        "eventos_b3.csv": texto_csv(COLUNAS_EVENTOS, juntar_eventos(
            ler_csv(pasta / "eventos_b3.csv"), eventos, resumo, excluidos)),
        "selic_mensal.csv": texto_csv(("mes", "taxa"), juntar_selic(
            ler_csv(pasta / "selic_mensal.csv"), ler_selic(coleta.selic, coleta.data), resumo)),
        "proventos_b3.csv": texto_csv(COLUNAS_PROVENTOS, juntar_proventos(
            ler_csv(pasta / "proventos_b3.csv"),
            [p for emissor, resp in sorted(coleta.eventos.items())
             for p in ler_proventos(emissor, resp)], resumo)),
        "proventos_historico_b3.csv": texto_csv(COLUNAS_HISTORICO, juntar_historico(
            ler_csv(pasta / "proventos_historico_b3.csv"),
            [p for emissor, resp in sorted(coleta.historico.items())
             for p in ler_historico(emissor, resp)], resumo)),
        "emissores_b3.csv": texto_csv(COLUNAS_EMISSORES, juntar_emissores(
            ler_csv(pasta / "emissores_b3.csv"), ler_emissores(coleta.empresas, coleta.fundos,
                                                               resumo), coleta.data, resumo)),
    }  # fmt: skip
    nomes = juntar_nomes(ler_csv(pasta / "nomes_pregao_b3.csv"),
                         [*coleta.nomes, *nomes_renda_fixa(coleta.fundos, coleta.data)],
                         resumo)  # fmt: skip
    arquivos["nomes_pregao_b3.csv"] = texto_csv(COLUNAS_NOMES, nomes)
    arquivos["tipos_ativo_b3.csv"] = texto_csv(("ticker", "tipo", "observacao"),
                                               tipos(historico, nomes))  # fmt: skip
    if (pasta / "trocas_ticker.csv").is_file():
        arquivos["trocas_ticker.csv"] = texto_csv(COLUNAS_TROCAS, juntar_trocas(
            ler_csv(pasta / "trocas_ticker.csv"), resumo.trocas_candidatas, historico,
            nomes))  # fmt: skip
    arquivos["pacote.json"] = json.dumps({"formato": FORMATO, "data": coleta.data.isoformat()},
                                         indent=1) + "\n"  # fmt: skip
    return arquivos, resumo


CURADOS = Path(__file__).with_name("eventos_cvm_ipe.csv")
EXCLUIDOS = Path(__file__).with_name("eventos_excluidos.csv")
COLUNAS_CHAVE = ("isin", "tipo", "data_com", "isin_emitido", "fator")  # chave do evento no pacote


def eventos_curados(caminho: Path = CURADOS) -> list[EventoB3]:
    """Eventos transcritos à mão dos avisos aos acionistas e fatos relevantes da CVM (IPE), cada um
    com o protocolo do documento: o que nenhuma fonte estruturada traz (2023 em diante)."""
    if not caminho.is_file():
        return []
    excluidos = eventos_excluidos()
    return [EventoB3(**{c: r.get(c, "") for c in COLUNAS_EVENTOS}) for r in ler_csv(caminho)
            if tuple(r[c] for c in COLUNAS_CHAVE) not in excluidos]  # fmt: skip


def eventos_excluidos(caminho: Path = EXCLUIDOS) -> set[tuple[str, ...]]:
    """Eventos retirados do pacote, cada um com o motivo e a fonte (``eventos_excluidos.csv``):
    nada sai sem registro, e o que saiu não volta na coleta seguinte."""
    if not caminho.is_file():
        return set()
    linhas = ler_csv(caminho)
    if any(not r["motivo"] or not r["fonte"] for r in linhas):
        raise GeracaoError(f"{caminho.name}: exclusão sem motivo ou sem fonte")
    return {tuple(r[c] for c in COLUNAS_CHAVE) for r in linhas}


def incorporar_eventos(pasta: Path, novos: Iterable[EventoB3],
                       data: date) -> tuple[dict[str, str], Resumo]:  # fmt: skip
    """``eventos_b3.csv`` com os eventos do histórico (FRE + COTAHIST) e os curados, e o resumo
    (novos e divergências). O ``pacote.json`` leva a data, se o arquivo mudou."""
    resumo = Resumo()
    texto = texto_csv(COLUNAS_EVENTOS, juntar_eventos(
        ler_csv(pasta / "eventos_b3.csv"), [*novos, *eventos_curados()], resumo,
        eventos_excluidos()))  # fmt: skip
    arquivos = {"eventos_b3.csv": texto,
                "pacote.json": json.dumps({"formato": FORMATO, "data": data.isoformat()},
                                          indent=1) + "\n"}  # fmt: skip
    return arquivos, resumo


def houve_pregao(dia: date, feriados: Iterable[Mapping[str, str]]) -> bool:
    """Fim de semana e os dias de ``feriados.csv`` (bancários e sem pregão) não têm pregão."""
    return dia.weekday() < 5 and dia.isoformat() not in {f["data"] for f in feriados}


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
