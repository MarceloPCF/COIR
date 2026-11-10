"""Histórico de desdobramentos, grupamentos e bonificações desde 2008 (o que a consulta de empresas
da B3 não traz: ela só guarda o último evento de cada tipo por ISIN).

Duas fontes oficiais, cruzadas, sem palpite:

- CVM, dados abertos do Formulário de Referência (FRE, 2010 a 2023), tabela
  ``capital_social_desdobramento`` (e ``..._classe_acao`` para PNA, PNB...): data de aprovação, tipo
  e quantidade de ações antes e depois, por espécie. O fator é a razão depois/antes, aceito só se
  for uma fração simples (p/q com q até 20, 1/n, n, ou percentual em meio ponto com tolerância
  menor): quantidades com ruído (frações vendidas
  em leilão, versões do formulário que não batem) vão para revisão;
- B3, série histórica de cotações (COTAHIST): o número de distribuição do ISIN muda no pregão
  seguinte à data com de todo evento (conferido em 100% dos eventos da consulta de empresas). A
  data com é o pregão anterior à única mudança, na janela depois da aprovação, em que o preço caiu
  (ou subiu) como o fator manda. Mudança ambígua, papel sem negócio no pregão da data ex ou fator
  pequeno demais para conferir pelo preço vão para revisão.

Puro e determinístico, só a biblioteca padrão (como o resto do gerador).
"""

from __future__ import annotations

import bisect
import csv
import io
import math
import re
import zipfile
from collections import defaultdict
from collections.abc import Collection, Iterable, Mapping
from dataclasses import dataclass, replace
from datetime import date, timedelta
from decimal import ROUND_CEILING, Decimal
from fractions import Fraction

from tools.referencias.gerar import EventoB3, multiplicador

DESDE = "2008-01-01"  # início do calendário de data/feriados.csv
TIPOS_FRE = {"Desd": "DESDOBRAMENTO", "Grup": "GRUPAMENTO", "Boni": "BONIFICACAO"}
CLASSES_FRE = {"Preferencial Classe A": "NPA", "Preferencial Classe B": "NPB",
               "Preferencial Classe C": "NPC"}  # fmt: skip
TOLERANCIA_FATOR = Fraction(2, 10000)  # frações de ação vendidas em leilão mexem no 5º dígito
TOLERANCIA_PERCENTUAL = Fraction(2, 100000)  # percentual em meio ponto: 1,0982 não vira 10%
JANELA = timedelta(days=270)  # da aprovação até a data ex (grupamento costuma levar meses)
PRECO_MINIMO = 0.02  # |ln fator| / 2 abaixo disto: a cotação do dia não distingue o evento


@dataclass(frozen=True)
class LinhaFre:
    cnpj: str  # só dígitos
    aprovado_em: str  # ISO
    tipo: str  # DESDOBRAMENTO, GRUPAMENTO, BONIFICACAO
    classe: str  # código da espécie no ISIN: NOR, NPR, NPA, NPB, NPC
    antes: int
    depois: int
    documento: str  # ID_Documento do FRE


def _inteiro(texto: str) -> int:
    try:
        return int(texto)
    except ValueError:
        return 0


def ler_fre(bruto: bytes) -> list[LinhaFre]:
    """Zip anual do FRE (dados.cvm.gov.br) -> linhas de desdobramento/grupamento/bonificação, por
    espécie. Zip sem a tabela (FRE de 2024 em diante) não tem linha."""
    with zipfile.ZipFile(io.BytesIO(bruto)) as z:
        nomes = z.namelist()
        base = next((n for n in nomes if "capital_social_desdobramento_2" in n), None)
        classes = next((n for n in nomes if "capital_social_desdobramento_classe" in n), None)
        if base is None:
            return []

        def ler(nome: str) -> list[list[str]]:
            return list(csv.reader(z.read(nome).decode("latin-1").splitlines()[1:], delimiter=";"))

        saida: list[LinhaFre] = []
        por_id: dict[str, tuple[str, str, str, str]] = {}
        for c in ler(base):
            tipo = TIPOS_FRE.get(c[7][:4])
            if tipo is None:
                continue
            chave = (re.sub(r"\D", "", c[0]), c[6], tipo, c[3])
            por_id[c[5]] = chave
            for classe, antes, depois in (("NOR", c[8], c[11]), ("NPR", c[9], c[12])):
                saida.append(LinhaFre(*chave[:3], classe, _inteiro(antes), _inteiro(depois),
                                      chave[3]))  # fmt: skip
        for c in ler(classes) if classes else []:
            chave_ = por_id.get(c[5])
            if chave_ is not None and c[6] in CLASSES_FRE:
                saida.append(LinhaFre(*chave_[:3], CLASSES_FRE[c[6]], _inteiro(c[7]),
                                      _inteiro(c[8]), chave_[3]))  # fmt: skip
    return saida


def ler_fca_codigos(bruto: bytes) -> dict[str, set[str]]:
    """Zip anual do FCA (Formulário Cadastral, dados.cvm.gov.br) -> CNPJ -> códigos de emissor (4
    primeiras letras dos códigos de negociação de ações e units). O FCA só traz o código de 2018 em
    diante; é o elo da companhia que saiu da bolsa (sem linha em emissores_b3.csv) com o FRE."""
    with zipfile.ZipFile(io.BytesIO(bruto)) as z:
        nome = next((n for n in z.namelist() if "valor_mobiliario" in n), None)
        if nome is None:
            return {}
        saida: dict[str, set[str]] = defaultdict(set)
        for r in csv.DictReader(z.read(nome).decode("latin-1").splitlines(), delimiter=";"):
            codigo = (r.get("Codigo_Negociacao") or "").strip().upper()
            if re.fullmatch(r"[A-Z0-9]{4}([3-8]|11)", codigo):
                saida[re.sub(r"\D", "", r["CNPJ_Companhia"])].add(codigo[:4])
    return dict(saida)


def deslistadas(codigos: Iterable[Mapping[str, set[str]]], listadas: Mapping[str, str],
                series: Series) -> dict[str, str]:  # fmt: skip
    """CNPJ -> código de emissor das companhias do FCA que não estão no cadastro de hoje (saíram da
    bolsa): o código com negócio mais recente no COTAHIST. CNPJ com mais de um código ainda
    negociado no mesmo último pregão fica de fora (ambíguo)."""
    ultimo: dict[str, str] = {}
    for isin, serie in series.items():
        if serie:
            ultimo[isin[2:6]] = max(ultimo.get(isin[2:6], ""), serie[-1].dia)
    juntos: dict[str, set[str]] = defaultdict(set)
    for ano in codigos:
        for cnpj, cods in ano.items():
            juntos[cnpj] |= cods
    saida: dict[str, str] = {}
    usados = set(listadas.values())
    for cnpj, cods in sorted(juntos.items()):
        vistos = sorted((ultimo[c], c) for c in cods if c in ultimo and c not in usados)
        if cnpj in listadas or not vistos:
            continue
        if len(vistos) > 1 and vistos[-1][0] == vistos[-2][0]:
            continue
        saida[cnpj] = vistos[-1][1]
    return saida


def fator_simples(antes: int, depois: int) -> Fraction | None:
    """depois/antes como fração simples (p/q com q <= 20, 1/n ou n) ou percentual em meio ponto
    (4% = 26/25, com tolerância bem menor: a grade é fina), ou None (ruído)."""
    if antes <= 0 or depois <= 0 or antes == depois:
        return None
    r = Fraction(depois, antes)
    candidatos = [(Fraction(round(r * q), q), TOLERANCIA_FATOR) for q in range(1, 21)]
    candidatos += [(Fraction(1, round(1 / r)), TOLERANCIA_FATOR)] if r < 1 else []
    if r > 1:
        candidatos.append((1 + Fraction(round((r - 1) * 200), 200), TOLERANCIA_PERCENTUAL))
    for f, tolerancia in candidatos:
        if f > 0 and abs(f / r - 1) < tolerancia:
            return f
    return None


def fator_b3(tipo: str, f: Fraction) -> str:
    """Como a B3 escreve: grupamento pela razão (1/30 -> 0,03333333334) e os demais em % (5/4 ->
    25). Dízima arredondada para cima, como a B3 (o piso da quantidade não perde uma ação)."""
    valor = f if tipo == "GRUPAMENTO" else (f - 1) * 100
    casas = 11 if tipo == "GRUPAMENTO" else 9
    d = Decimal(valor.numerator) / Decimal(valor.denominator)
    texto = format(d.quantize(Decimal(1).scaleb(-casas), rounding=ROUND_CEILING).normalize(), "f")
    return texto


# --------------------------------------------------------------------------------- COTAHIST
@dataclass(frozen=True)
class Pregao:
    dia: str  # AAAAMMDD
    distribuicao: str
    preco: float  # último preço por ação (já dividido pelo fator de cotação)
    isin: str = ""  # ISIN negociado no pregão, se não for o da série (código anterior a uma troca)


Series = dict[str, list[Pregao]]  # ISIN -> pregões em que houve negócio (lote padrão, à vista)


def pares_de_troca(trocas: Iterable[Mapping[str, str]]) -> list[tuple[str, str]]:
    """``trocas_ticker.csv`` -> (ISIN antigo, ISIN novo) das trocas de código que mudaram o ISIN
    (a troca tratada como incorporação, sem ISIN antigo, fica de fora)."""
    return [(t["isin_antigo"], t["isin_novo"]) for t in trocas
            if t.get("isin_antigo") and t.get("isin_novo")
            and t["isin_antigo"] != t["isin_novo"]]  # fmt: skip


def encadear(series: Series, pares: Iterable[tuple[str, str]]) -> Series:
    """Troca de código (KROT3 -> COGN3, GLOB3 -> VVAR3 -> VIIA3 -> BHIA3): o ISIN muda junto, mas
    o papel é o mesmo. A série do ISIN mais novo passa a começar pelos pregões dos antigos (cada um
    marcado com o ISIN em que foi negociado), e os antigos saem como chave: o evento de antes da
    troca deixa de se perder porque o código atual não bate, e o que tem a data ex no dia da troca
    (grupamento 8:1 de LLIS3/VSTE3, data ex 09/02/2023) aparece como a variação entre os dois.
    O par cujo ISIN antigo ainda negocia depois do primeiro pregão do novo (o código voltou, como
    CTAX3 -> LIQO3 -> CTAX3) não é encadeado: as séries se misturariam."""
    seguinte = {a: n for a, n in pares
                if not (series.get(a) and series.get(n) and series[a][-1].dia >= series[n][0].dia)
                }  # fmt: skip
    anterior = {n: a for a, n in seguinte.items()}
    saida = dict(series)
    for novo in sorted(n for n in anterior if n not in seguinte):
        cadeia = [novo]
        while cadeia[-1] in anterior and len(cadeia) <= len(anterior):
            cadeia.append(anterior[cadeia[-1]])
        pregoes: list[Pregao] = []
        for isin in reversed(cadeia[1:]):
            pregoes += [replace(p, isin=isin) for p in saida.pop(isin, [])]
        if pregoes:
            saida[novo] = sorted(pregoes + saida.get(novo, []), key=lambda p: p.dia)
    return saida


def negociado(serie: list[Pregao], com: str, chave: str) -> str:
    """ISIN negociado no pregão seguinte à data com (a data ex): quem tinha o papel na data com
    recebe o evento no código que vale na data ex."""
    j = bisect.bisect_right([p.dia for p in serie], com.replace("-", ""))
    return (serie[j].isin or chave) if j < len(serie) else chave


def ler_cotahist(bruto: bytes, series: Series, pregoes: set[str],
                 emissores: set[str] | None = None) -> None:  # fmt: skip
    """Zip anual do COTAHIST (layout posicional da B3) -> acrescenta em ``series`` os pregões de
    ações e units (ISIN com "AC" ou "CD") do mercado à vista, lote padrão; ``pregoes`` recebe
    todos os dias com negócio (o calendário da bolsa)."""
    with zipfile.ZipFile(io.BytesIO(bruto)) as z, z.open(z.namelist()[0]) as f:
        for linha in f:
            if linha[:2] != b"01":
                continue
            dia = linha[2:10].decode()
            pregoes.add(dia)
            if linha[24:27] != b"010" or linha[10:12] != b"02":
                continue
            isin = linha[230:242].decode()
            if isin[6:8] not in ("AC", "CD") or (emissores is not None
                                                 and isin[2:6] not in emissores):  # fmt: skip
                continue
            preco = int(linha[108:121]) / 100 / int(linha[210:217])
            if preco > 0:
                series.setdefault(isin, []).append(Pregao(dia, linha[242:245].decode(), preco))


def nome_de_pregao(nome: str, especificacao: str, sufixos: Collection[str]) -> str:
    """Nome + especificação como a nota imprime, sem os sufixos do fim que mudam com o tempo ("ON
    ED NM" -> "ON"); a leitura da nota tira os mesmos. Os de BDR (DR1, DRN...) ficam: separam dois
    programas do mesmo nome."""
    tokens = especificacao.split()
    while tokens and tokens[-1] in sufixos and not tokens[-1].startswith("DR"):
        tokens.pop()
    return " ".join([*nome.split(), *tokens])


RE_TICKER_11 = re.compile(r"^[A-Z0-9]{4}11$")
BDI_FUNDOS = (b"12", b"13")  # códigos BDI do COTAHIST: fundos imobiliários e Fiagro


def _tipo_no_cotahist(ticker: str, bdi: bytes, especificacao: str) -> str:
    """Tipo que o próprio COTAHIST dá ao código final 11 (o resto vem do final do ticker)."""
    if not RE_TICKER_11.match(ticker):
        return ""
    classe = (especificacao.split() or [""])[0]
    if bdi in BDI_FUNDOS and classe == "CI":
        return "FII"
    if classe == "UNT":
        return "UNIT"
    return "BDR" if classe.startswith("DR") else ""


def ler_nomes(bruto: bytes, sufixos: Collection[str]) -> list[dict[str, str]]:
    """Zip anual do COTAHIST -> nome de pregão, ticker, período e tipo dos negócios à vista (lote
    padrão, "010", e fracionário, "020", com o F tirado). É o que a nota imprime, ano a ano."""
    periodos: dict[tuple[str, str], list[str]] = {}
    with zipfile.ZipFile(io.BytesIO(bruto)) as z, z.open(z.namelist()[0]) as f:
        for linha in f:
            if linha[:2] != b"01" or linha[24:27] not in (b"010", b"020"):
                continue
            ticker = linha[12:24].decode("latin-1").strip()
            if linha[24:27] == b"020" and ticker.endswith("F"):
                ticker = ticker[:-1]
            especificacao = linha[39:49].decode("latin-1")
            nome = nome_de_pregao(linha[27:39].decode("latin-1"), especificacao, sufixos)
            dia = linha[2:10].decode()
            periodo = periodos.setdefault(
                (nome, ticker), [dia, dia, _tipo_no_cotahist(ticker, linha[10:12], especificacao)]
            )
            periodo[0], periodo[1] = min(periodo[0], dia), max(periodo[1], dia)
    return [{"nome_pregao": n, "ticker": t, "desde": _iso(p[0]), "ate": _iso(p[1]), "tipo": p[2]}
            for (n, t), p in sorted(periodos.items())]  # fmt: skip


def data_com(serie: list[Pregao], pregoes: list[str], aprovado_em: str, fator: Fraction,
             ate: str = "") -> tuple[str | None, str]:  # fmt: skip
    """(data com ISO, "") ou (None, motivo da revisão). ``serie`` em ordem; ``ate`` limita a janela
    (a próxima aprovação do mesmo tipo e espécie, para não pegar o evento do ano seguinte)."""
    dias = [p.dia for p in serie]
    a = aprovado_em.replace("-", "")
    limite = (date.fromisoformat(aprovado_em) + JANELA).strftime("%Y%m%d")
    if ate:
        limite = min(limite, ate.replace("-", ""))
    i = bisect.bisect_right(dias, a)
    if i == 0:
        return None, "antes da listagem"
    if i >= len(serie):
        return None, "sem negócio depois da aprovação"
    esperado = -math.log(fator)
    tolerancia = abs(esperado) / 2
    if tolerancia < PRECO_MINIMO:
        return None, "fator pequeno demais para conferir pela cotação"
    mudancas: list[tuple[float, bool, str, str]] = []  # (desvio, com negócio na véspera, com, ex)
    for j in range(i, len(serie)):
        if serie[j].dia > limite:
            break
        if serie[j].distribuicao == serie[j - 1].distribuicao:
            continue
        k = bisect.bisect_left(pregoes, serie[j].dia)
        vespera = pregoes[k - 1]
        desvio = abs(math.log(serie[j].preco / serie[j - 1].preco) - esperado)
        mudancas.append((desvio, vespera == serie[j - 1].dia, vespera, serie[j].dia))
    coerentes = [m for m in mudancas if m[0] < tolerancia]
    if not coerentes:
        return None, "nenhuma mudança de distribuição com a cotação coerente com o fator"
    if len(coerentes) > 1:
        return None, "mais de uma data ex possível: " + ", ".join(_iso(m[3]) for m in coerentes)
    _, _, com, ex = coerentes[0]
    if not all(m[1] for m in mudancas if m[3] <= ex):
        return None, f"papel sem negócio em algum pregão até a data ex ({_iso(ex)})"
    return _iso(com), ""


def _iso(dia: str) -> str:
    return f"{dia[:4]}-{dia[4:6]}-{dia[6:]}"


# --------------------------------------------------------------------------------- junção
@dataclass(frozen=True)
class Revisao:
    emissor: str
    isin: str
    tipo: str
    aprovado_em: str
    fator: str  # razão depois/antes (ou as razões das versões, se divergem)
    motivo: str
    documento: str
    razao: float = 0.0  # depois/antes da primeira versão (para achar o evento em outra fonte)

    def texto(self) -> str:
        return (f"{self.emissor} {self.isin} {self.tipo} aprovado em {self.aprovado_em}, "
                f"fator {self.fator}: {self.motivo} (FRE {self.documento})")  # fmt: skip


IGNORAR = ("antes da listagem",)  # evento antes do primeiro negócio: não atinge investidor da B3


def eventos_fre(linhas: Iterable[LinhaFre], emissor_por_cnpj: Mapping[str, str],
                series: Series, pregoes: Iterable[str],
                desde: str = DESDE) -> tuple[list[EventoB3], list[Revisao]]:  # fmt: skip
    """Eventos do FRE com data com conferida no COTAHIST, e os que vão para revisão humana.
    Só companhias do cadastro (CNPJ em emissores_b3.csv) e ISIN com negócio no período."""
    calendario = sorted(pregoes)
    grupos: dict[tuple[str, str, str, str], list[LinhaFre]] = defaultdict(list)
    for x in linhas:
        if x.aprovado_em >= desde and x.cnpj in emissor_por_cnpj and x.antes != x.depois:
            grupos[(x.cnpj, x.tipo, x.classe, x.aprovado_em)].append(x)
    aprovacoes: dict[tuple[str, str, str], list[str]] = defaultdict(list)
    for cnpj, tipo, classe, aprovado in grupos:
        aprovacoes[(cnpj, tipo, classe)].append(aprovado)
    no_mesmo_ato = {(c, k, d) for c, t, k, d in grupos if t == "GRUPAMENTO"
                    and (c, "DESDOBRAMENTO", k, d) in grupos}  # fmt: skip
    eventos: list[EventoB3] = []
    revisao: list[Revisao] = []
    pendentes: list[tuple[str, Revisao]] = []  # sem data com pela cotação desta espécie
    confirmadas: dict[tuple[str, str, str, str], str] = {}  # (cnpj, tipo, aprovação, fator) -> com
    for (cnpj, tipo, classe, aprovado), versoes in sorted(grupos.items()):
        emissor = emissor_por_cnpj[cnpj]
        isins = sorted(i for i in series if i[2:6] == emissor and i[6:8] == "AC"
                       and i[8:11] == classe)  # fmt: skip
        documento = max(v.documento for v in versoes)
        validas = [v for v in versoes if v.antes > 0 and v.depois > 0]
        if not isins or not validas:
            continue  # espécie sem negócio (ou sem quantidade): nada a ajustar na B3
        fatores = {fator_simples(v.antes, v.depois) for v in validas}
        seguintes = [d for d in sorted(aprovacoes[(cnpj, tipo, classe)]) if d > aprovado]
        razao = validas[0].depois / validas[0].antes
        for isin in isins:
            if (cnpj, classe, aprovado) in no_mesmo_ato:
                razoes = ", ".join(sorted({f"{v.depois}/{v.antes}" for v in validas}))
                revisao.append(Revisao(emissor, isin, tipo, aprovado, razoes,
                                       "grupamento e desdobramento no mesmo ato: a cotação não "
                                       "mostra a data ex", documento, razao))  # fmt: skip
                continue
            if len(fatores) != 1 or None in fatores:
                razoes = ", ".join(sorted({f"{v.depois}/{v.antes}" for v in validas}))
                motivo = ("versões do FRE com fatores diferentes" if len(fatores) > 1
                          else "fator não é uma fração simples")  # fmt: skip
                revisao.append(Revisao(emissor, isin, tipo, aprovado, razoes, motivo, documento,
                                       razao))  # fmt: skip
                continue
            f = next(iter(fatores))
            assert f is not None
            if (tipo == "GRUPAMENTO") != (f < 1):
                motivo = "fator incoerente com o tipo"
                revisao.append(Revisao(emissor, isin, tipo, aprovado, str(f), motivo, documento,
                                       razao))  # fmt: skip
                continue
            com, motivo = data_com(series[isin], calendario, aprovado, f,
                                   seguintes[0] if seguintes else "")  # fmt: skip
            if com is None:
                if motivo not in IGNORAR:
                    pendentes.append((cnpj, Revisao(emissor, isin, tipo, aprovado, str(f), motivo,
                                                    documento, razao)))  # fmt: skip
                continue
            confirmadas[(cnpj, tipo, aprovado, str(f))] = com
            ex = negociado(series[isin], com, isin)
            eventos.append(EventoB3(
                emissor=emissor, isin=ex, tipo=tipo, fator=fator_b3(tipo, f), data_com=com,
                aprovado_em=aprovado, isin_emitido=ex, origem="cvm_fre",
                referencia=f"FRE {documento}",
            ))  # fmt: skip
    for cnpj, r in pendentes:
        com_irma = confirmadas.get((cnpj, r.tipo, r.aprovado_em, r.fator))
        mudou = com_irma is not None and r.motivo.startswith(PELA_IRMA) and (
            muda_no_pregao_seguinte(series[r.isin], calendario, com_irma) is not None)  # fmt: skip
        if com_irma is None or not mudou:
            revisao.append(r)
            continue
        ex = negociado(series[r.isin], com_irma, r.isin)
        eventos.append(EventoB3(
            emissor=r.emissor, isin=ex, tipo=r.tipo, fator=fator_b3(r.tipo, Fraction(r.fator)),
            data_com=com_irma, aprovado_em=r.aprovado_em, isin_emitido=ex, origem="cvm_fre",
            referencia=f"FRE {r.documento} (data com da outra espécie)",
        ))  # fmt: skip
    # na revisão, o ISIN que valia na aprovação (KROT, não COGN, para o ato de 2012)
    revisao = [replace(r, isin=negociado(series[r.isin], r.aprovado_em, r.isin))
               if r.isin in series else r for r in revisao]  # fmt: skip
    eventos, ambiguos = _sem_tipo_ambiguo(eventos)
    revisao += ambiguos
    eventos += _units(eventos, series, calendario)
    return eventos, revisao


def _sem_tipo_ambiguo(eventos: list[EventoB3]) -> tuple[list[EventoB3], list[Revisao]]:
    """Versões do FRE que chamam o mesmo ato (ISIN, data com e fator) de bonificação numa e de
    desdobramento noutra: as duas linhas dobrariam a quantidade, e o tipo muda o custo. Nenhuma
    entra; o ato vai para revisão (o aviso aos acionistas decide)."""
    por_ato: dict[tuple[str, str, Fraction], set[str]] = defaultdict(set)
    for e in eventos:
        if e.tipo in AUMENTAM:
            por_ato[(e.isin, e.data_com, multiplicador(e.tipo, e.fator))].add(e.tipo)
    ambiguos = {k for k, tipos in por_ato.items() if len(tipos) > 1}
    revisao = [
        Revisao(e.emissor, e.isin, e.tipo, e.aprovado_em, e.fator,
                "versões do FRE com tipos diferentes (bonificação e desdobramento)",
                e.referencia.removeprefix("FRE ").split()[0], float(multiplicador(e.tipo, e.fator)))
        for e in eventos
        if e.tipo in AUMENTAM and (e.isin, e.data_com, multiplicador(e.tipo, e.fator)) in ambiguos
    ]  # fmt: skip
    restantes = [e for e in eventos if not (e.tipo in AUMENTAM and (
        e.isin, e.data_com, multiplicador(e.tipo, e.fator)) in ambiguos)]  # fmt: skip
    return restantes, revisao


AUMENTAM = ("DESDOBRAMENTO", "BONIFICACAO")  # mesmo efeito na quantidade: o FRE confunde os dois


def _units(eventos: list[EventoB3], series: Series, calendario: list[str]) -> list[EventoB3]:
    """Unit do emissor (ISIN com "CD") no mesmo evento, quando todas as espécies com negócio tiveram
    o mesmo fator e a mesma data com e a unit também mudou de distribuição no pregão seguinte."""
    por_ato: dict[tuple[str, str, str], list[EventoB3]] = defaultdict(list)
    for e in eventos:
        por_ato[(e.emissor, e.tipo, e.aprovado_em)].append(e)
    saida = []
    for (emissor, tipo, _), evs in sorted(por_ato.items()):
        if len({(e.fator, e.data_com) for e in evs}) != 1:
            continue
        e = evs[0]
        esperado = -math.log(multiplicador(tipo, e.fator))
        for unit in sorted(i for i in series if i[2:6] == emissor and i[6:8] == "CD"):
            variacao = muda_no_pregao_seguinte(series[unit], calendario, e.data_com)
            if variacao is not None and abs(variacao - esperado) < abs(esperado) / 2:
                referencia = e.referencia + " (unit: o fator das ações)"
                ex = negociado(series[unit], e.data_com, unit)
                saida.append(replace(e, isin=ex, isin_emitido=ex, referencia=referencia))
    return saida


# a cotação desta espécie não confirmou, mas outra espécie do mesmo ato confirmou a data com
PELA_IRMA = ("nenhuma mudança", "mais de uma", "fator pequeno")


def muda_no_pregao_seguinte(serie: list[Pregao], pregoes: list[str], com: str) -> float | None:
    """Variação (ln) do preço da data com para o pregão seguinte, se houve negócio nos dois e o
    número de distribuição mudou entre eles (a data ex da outra espécie vale para esta); senão
    None."""
    dia = com.replace("-", "")
    dias = [p.dia for p in serie]
    i = bisect.bisect_left(dias, dia)
    k = bisect.bisect_right(pregoes, dia)
    if (i + 1 < len(serie) and dias[i] == dia and k < len(pregoes) and dias[i + 1] == pregoes[k]
            and serie[i + 1].distribuicao != serie[i].distribuicao):  # fmt: skip
        return math.log(serie[i + 1].preco / serie[i].preco)
    return None


COBERTURA = 0.005  # o FRE arredonda as quantidades; a B3 traz o fator exato


def cobertas(revisao: Iterable[Revisao], linhas: Iterable[Mapping[str, str]]) -> list[Revisao]:
    """Tira da revisão o que outra fonte já traz: mesma espécie e tipo, data com na janela depois
    da aprovação e fator igual ao do FRE a menos do arredondamento das quantidades."""
    por_isin: dict[tuple[str, str], list[Mapping[str, str]]] = defaultdict(list)
    for r in linhas:
        por_isin[(r["isin"], r["tipo"])].append(r)
    saida = []
    for x in revisao:
        fim = (date.fromisoformat(x.aprovado_em) + JANELA).isoformat()
        if not any(x.aprovado_em <= r["data_com"] <= fim
                   and abs(float(multiplicador(r["tipo"], r["fator"])) / x.razao - 1) < COBERTURA
                   for r in por_isin[(x.isin, x.tipo)]):  # fmt: skip
            saida.append(x)
    return saida
