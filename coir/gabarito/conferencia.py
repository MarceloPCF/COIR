"""Conferência automática gabarito x PDFs das notas, em todas as notas de todos os PDFs.

Para cada nota do gabarito, localiza as páginas do PDF com o número da nota e confere:
- Bovespa: por C/V, soma das quantidades e dos valores das linhas de negócio; líquido da nota
  ("Líquido para") contra vendas - compras - custos (com ou sem o IRRF descontado);
- BM&F: por C/V, soma das quantidades e de quantidade x preço; em notas só de day trade,
  ajuste day trade, IRRF e líquido esperados presentes entre os valores impressos na página.
Também lista páginas com negócios que não pertencem a nenhuma nota do gabarito e notas do
gabarito que não estão em nenhum PDF.

É uma ferramenta de validação do gabarito (F1), não o extrator da F2: lê só o necessário.
"""

import re
from collections import Counter, defaultdict
from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path

from coir.erros import GabaritoError
from coir.gabarito.modelo import Gabarito, OperacaoGabarito
from coir.gabarito.piloto import (
    MERCADOS_BMF,
    ArquivoNota,
    interpretar_nome,
    liquido_pdf,
    paginas_da_nota,
)
from coir.parsing.texto_pdf import PdfError, paginas

UM_CENTAVO = Decimal("0.01")
RE_NUMERO = re.compile(r"(?<![\w/,.])(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d+)?(?![\w/])")
RE_DINHEIRO = re.compile(r"(?<![\w/,.])(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2}(?![\d\w/])")
RE_BMF = re.compile(
    r"^(?P<cv>[CV])\s+(?P<contrato>.+?)\s+(?:@?\d{2}/\d{2}/\d{4}\s+)?(?P<qtd>\d[\d.]*)\s+"
    r"(?P<preco>[\d.]+,\d+)\s+(?P<tipo>.+?)\s+(?P<valor>[\d.]+,\d{2})(?:\s+(?P<dc>[DC]))?(?:\s|$)"
)  # D/C ausente em linhas de abertura com valor 0,00 (XP BM&F 2024+)
TIPOS_BMF_IGNORADOS = ("AJUPOS", "PERMAN")


def numero(texto: str) -> Decimal:
    return Decimal(texto.replace(".", "").replace(",", "."))


@dataclass(frozen=True)
class LinhaNegocio:
    cv: str
    quantidade: Decimal
    preco: Decimal
    valor: Decimal


def linhas_bovespa(texto: str) -> list[LinhaNegocio]:
    """Linhas de negócio Bovespa: '... BOVESPA C|V ... quantidade preço [D|C] valor [D|C]'."""
    saida: list[LinhaNegocio] = []
    for linha in texto.splitlines():
        m = re.search(r"BOVESPA\s+(?P<cv>[CV])\s", linha)
        if not m:
            continue
        tokens = RE_NUMERO.findall(linha[m.end() :])
        if len(tokens) < 3 or "," not in tokens[-1] or "," not in tokens[-2] or "," in tokens[-3]:
            continue
        saida.append(
            LinhaNegocio(m["cv"], numero(tokens[-3]), numero(tokens[-2]), numero(tokens[-1]))
        )
    return saida


def linhas_bmf(texto: str) -> list[LinhaNegocio]:
    """Linhas de negócio BM&F (exclui ajuste de posição e taxa de permanência)."""
    saida: list[LinhaNegocio] = []
    for linha in texto.splitlines():
        m = RE_BMF.match(linha.strip())
        if not m or any(t in m["tipo"].upper() for t in TIPOS_BMF_IGNORADOS):
            continue
        quantidade = numero(m["qtd"])
        if quantidade == 0:
            continue
        saida.append(LinhaNegocio(m["cv"], quantidade, numero(m["preco"]), numero(m["valor"])))
    return saida


def valores_impressos(texto: str) -> set[Decimal]:
    return {numero(t) for t in RE_DINHEIRO.findall(texto)}


def _presente(valor: Decimal, impressos: set[Decimal]) -> bool:
    alvo = abs(valor).quantize(UM_CENTAVO)
    return any(abs(alvo - v) <= UM_CENTAVO for v in impressos)


@dataclass
class Verificacao:
    nome: str
    esperado: Decimal | None
    lido: Decimal | None
    ok: bool | None  # None = não verificável neste layout
    observacao: str = ""


@dataclass
class ResultadoNota:
    arquivo: str
    nota: str
    data: str
    conta: str
    verificacoes: list[Verificacao] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return all(v.ok is not False for v in self.verificacoes)


@dataclass
class ResultadoArquivo:
    arquivo: str
    notas: list[ResultadoNota] = field(default_factory=list)
    paginas_sem_nota: list[int] = field(default_factory=list)  # páginas (1..n) com negócios órfãos
    erro: str = ""  # PDF não lido (nome fora do padrão, ilegível)
    ausente: bool = False  # PDF esperado pelas notas do gabarito, mas inexistente


def _comparar(
    nome: str, esperado: Decimal, lido: Decimal, tol: Decimal = UM_CENTAVO
) -> Verificacao:
    return Verificacao(nome, esperado, lido, abs(esperado - lido) <= tol)


def _somas(linhas: Iterable[LinhaNegocio], bmf: bool) -> dict[str, Decimal]:
    s: dict[str, Decimal] = defaultdict(Decimal)
    for linha in linhas:
        s[f"quantidade {linha.cv}"] += linha.quantidade
        chave = "quantidade x preço" if bmf else "valor"
        s[f"{chave} {linha.cv}"] += linha.quantidade * linha.preco if bmf else linha.valor
    return s


def _somas_gabarito(ops: list[OperacaoGabarito], bmf: bool) -> dict[str, Decimal]:
    s: dict[str, Decimal] = defaultdict(Decimal)
    for op in ops:
        s[f"quantidade {op.cv}"] += op.quantidade
        chave = "quantidade x preço" if bmf else "valor"
        s[f"{chave} {op.cv}"] += op.quantidade * op.preco if bmf else op.total
    return s


def conferir_nota(
    arquivo: str, ops: list[OperacaoGabarito], texto: str, bmf: bool
) -> ResultadoNota:
    res = ResultadoNota(arquivo, ops[0].nota, ops[0].data.isoformat(), ops[0].conta)
    linhas = linhas_bmf(texto) if bmf else linhas_bovespa(texto)
    lidas, esperadas = _somas(linhas, bmf), _somas_gabarito(ops, bmf)
    for chave in sorted(esperadas.keys() | lidas.keys()):
        # quantidade x preço (BM&F) é em pontos/preço do contrato: tolerância de 1 unidade
        tol = Decimal(1) if chave.startswith("quantidade x") else UM_CENTAVO
        res.verificacoes.append(
            _comparar(chave, esperadas.get(chave, Decimal(0)), lidas.get(chave, Decimal(0)), tol)
        )
    compras = sum((o.total for o in ops if o.cv == "C"), Decimal(0))
    vendas = sum((o.total for o in ops if o.cv == "V"), Decimal(0))
    custos = sum((o.custos for o in ops), Decimal(0))
    irrf = sum((o.irrf for o in ops), Decimal(0))
    estimado = vendas - compras - custos
    if not bmf:
        liquido = liquido_pdf(texto)
        if liquido is None:
            res.verificacoes.append(Verificacao("líquido", estimado, None, None, "não impresso"))
        elif abs(liquido - estimado) <= UM_CENTAVO:
            res.verificacoes.append(Verificacao("líquido", estimado, liquido, True))
        else:
            v = _comparar("líquido", estimado - irrf, liquido)
            v.observacao = "com IRRF descontado" if v.ok else ""
            res.verificacoes.append(v)
        return res
    qtd: dict[tuple[str, str], Decimal] = defaultdict(Decimal)
    for op in ops:
        qtd[(op.papel, op.cv)] += op.quantidade
    papeis = {p for p, _ in qtd}
    so_day_trade = all(qtd.get((p, "C"), 0) == qtd.get((p, "V"), 0) for p in papeis)
    if not so_day_trade:
        res.verificacoes.append(
            Verificacao("líquido", None, None, None, "posição em aberto: depende de ajuste")
        )
        return res
    impressos = valores_impressos(texto)
    ajuste = vendas - compras
    for nome, valor in (
        ("ajuste day trade", ajuste),
        ("IRRF day trade", irrf),
        ("líquido", ajuste - custos - irrf),
    ):
        if nome == "IRRF day trade" and not valor:
            continue
        if not valor:  # 0,00 aparece em quase toda nota: não prova nada
            res.verificacoes.append(Verificacao(nome, valor, None, None, "zero: não verificável"))
            continue
        res.verificacoes.append(
            Verificacao(nome, valor, None, _presente(valor, impressos), "procurado no resumo")
        )
    return res


def _notas_do_gabarito(gabarito: Gabarito) -> dict[ArquivoNota, list[list[OperacaoGabarito]]]:
    """Operações agrupadas pelo PDF esperado (corretora, mês do pregão, segmento) e pela nota
    (conta + número + data: a numeração BM&F da XP se repete no mesmo mês)."""
    grupos: dict[ArquivoNota, dict[tuple[str, str, date], list[OperacaoGabarito]]] = defaultdict(
        lambda: defaultdict(list)
    )
    for op in gabarito.operacoes:
        arq = ArquivoNota(op.corretora, op.data.year, op.data.month, op.mercado in MERCADOS_BMF)
        grupos[arq][(op.conta, op.nota, op.data)].append(op)
    return {arq: list(notas.values()) for arq, notas in grupos.items()}


def _paginas(textos: list[str], livres: list[int], ops: list[OperacaoGabarito]) -> list[int]:
    """Páginas livres com o número da nota; havendo mais de uma, só as com a data do pregão
    (dd/mm/aaaa ou dd/mm/aa)."""
    pags = [livres[i] for i in paginas_da_nota([textos[j] for j in livres], ops[0].nota)]
    datas = (ops[0].data.strftime("%d/%m/%Y"), ops[0].data.strftime("%d/%m/%y"))
    com_data = [i for i in pags if any(d in textos[i] for d in datas)]
    return com_data if len(pags) > 1 and com_data else pags


def _nao_encontrada(arquivo: str, ops: list[OperacaoGabarito], obs: str) -> ResultadoNota:
    r = ResultadoNota(arquivo, ops[0].nota, ops[0].data.isoformat(), ops[0].conta)
    r.verificacoes.append(Verificacao("nota no PDF", None, None, False, obs))
    return r


def conferir_textos(
    gabarito: Gabarito, textos_por_arquivo: dict[str, list[str]]
) -> list[ResultadoArquivo]:
    """Confere todas as notas do gabarito contra os textos (por página) dos PDFs.

    Cada nota é procurada primeiro no PDF do seu segmento (Corretora_NC_AAAA_MM[_BMF].pdf); se
    não estiver lá, nas páginas ainda não usadas dos outros PDFs da corretora no mês (há PDFs
    mistos, com ações e BM&F no mesmo arquivo). Nota sem PDF nenhum também vira divergência.
    """
    arquivos = {nome: interpretar_nome(nome) for nome in sorted(textos_por_arquivo)}
    por_arquivo = {nome: ResultadoArquivo(nome) for nome in arquivos}
    usadas: dict[str, set[int]] = {nome: set() for nome in arquivos}
    por_mes: dict[tuple[str, int, int], list[str]] = defaultdict(list)
    for nome, arq in arquivos.items():
        por_mes[(arq.corretora, arq.ano, arq.mes)].append(nome)
    preferido = {arq: nome for nome, arq in arquivos.items()}

    def conferir_em(nome: str, ops: list[OperacaoGabarito], pags: list[int]) -> None:
        usadas[nome].update(pags)
        texto = "\n".join(textos_por_arquivo[nome][i] for i in pags)
        por_arquivo[nome].notas.append(
            conferir_nota(nome, ops, texto, ops[0].mercado in MERCADOS_BMF)
        )

    def livres(nome: str) -> list[int]:
        return [i for i in range(len(textos_por_arquivo[nome])) if i not in usadas[nome]]

    pendentes: list[tuple[ArquivoNota, list[OperacaoGabarito]]] = []
    for arq, notas in _notas_do_gabarito(gabarito).items():
        nome = preferido.get(arq)
        repetidas = Counter((ops[0].nota, ops[0].data) for ops in notas)
        for ops in notas:
            if nome and repetidas[(ops[0].nota, ops[0].data)] > 1:
                # sem a conta real (pseudonimizada) não há como separar as páginas
                r = ResultadoNota(nome, ops[0].nota, ops[0].data.isoformat(), ops[0].conta)
                obs = "mesmo número e data em outra conta"
                r.verificacoes.append(Verificacao("nota no PDF", None, None, None, obs))
                por_arquivo[nome].notas.append(r)
                continue
            # no PDF do próprio segmento todas as páginas valem (como antes); só a busca
            # nos outros PDFs do mês fica restrita às páginas que nenhuma nota usou
            todas = list(range(len(textos_por_arquivo[nome]))) if nome else []
            pags = _paginas(textos_por_arquivo[nome], todas, ops) if nome else []
            if nome and pags:
                conferir_em(nome, ops, pags)
            else:
                pendentes.append((arq, ops))
    sem_pdf: dict[str, ResultadoArquivo] = {}
    for arq, ops in pendentes:
        for nome in por_mes.get((arq.corretora, arq.ano, arq.mes), []):
            pags = _paginas(textos_por_arquivo[nome], livres(nome), ops)
            if pags:
                conferir_em(nome, ops, pags)
                break
        else:
            nome = preferido.get(arq)
            if nome:
                por_arquivo[nome].notas.append(_nao_encontrada(nome, ops, "não encontrada"))
            else:
                esperado = (
                    f"{arq.corretora}_NC_{arq.ano}_{arq.mes:02d}{'_BMF' if arq.bmf else ''}.pdf"
                )
                res = sem_pdf.setdefault(
                    esperado, ResultadoArquivo(f"{esperado} (ausente)", ausente=True)
                )
                res.notas.append(_nao_encontrada(res.arquivo, ops, "PDF ausente"))
    for nome, res in por_arquivo.items():
        res.notas.sort(key=lambda n: (n.data, n.nota))
        textos = textos_por_arquivo[nome]
        res.paginas_sem_nota = [
            i + 1
            for i, t in enumerate(textos)
            if i not in usadas[nome] and (linhas_bovespa(t) or linhas_bmf(t))
        ]
    return [*por_arquivo.values(), *(sem_pdf[k] for k in sorted(sem_pdf))]


def conferir_pasta(gabarito: Gabarito, pasta: Path) -> list[ResultadoArquivo]:
    """Lê os PDFs da pasta e confere; PDF fora do padrão de nome ou ilegível entra no relatório
    com o motivo (as notas dele aparecem como "PDF ausente")."""
    textos: dict[str, list[str]] = {}
    fora: list[ResultadoArquivo] = []
    for caminho in sorted(pasta.glob("*.pdf")):
        try:
            interpretar_nome(caminho.name)
            textos[caminho.name] = paginas(caminho)
        except (GabaritoError, PdfError) as erro:  # ex.: opção flexível, fora do escopo
            fora.append(ResultadoArquivo(caminho.name, erro=str(erro)))
    return [*conferir_textos(gabarito, textos), *fora]


def _fmt(valor: Decimal | None) -> str:
    return "" if valor is None else format(valor.quantize(UM_CENTAVO), "f")


def relatorio(resultados: list[ResultadoArquivo]) -> tuple[str, list[list[str]]]:
    """Markdown com o resumo e as divergências + linhas de CSV com todas as verificações."""
    csv: list[list[str]] = [["arquivo", "nota", "data", "conta", "verificacao", "esperado", "lido",
                             "ok", "observacao"]]  # fmt: skip
    total_notas = notas_ok = verif = verif_ok = nao_verif = 0
    divergentes: list[str] = []
    for arq in resultados:
        for nota in arq.notas:
            total_notas += 1
            notas_ok += nota.ok
            for v in nota.verificacoes:
                verif += 1
                verif_ok += v.ok is True
                nao_verif += v.ok is None
                situacao = {True: "sim", False: "não", None: "não verificado"}[v.ok]
                csv.append([arq.arquivo, nota.nota, nota.data, nota.conta, v.nome,
                            _fmt(v.esperado), _fmt(v.lido), situacao, v.observacao])  # fmt: skip
                if v.ok is False:
                    divergentes.append(
                        f"| {arq.arquivo} | {nota.nota} | {nota.data} | {v.nome} "
                        f"| {_fmt(v.esperado)} | {_fmt(v.lido)} | {v.observacao} |"
                    )
    orfas = [
        f"- {a.arquivo}: páginas {a.paginas_sem_nota}" for a in resultados if a.paginas_sem_nota
    ]
    nao_lidos = [f"- {a.arquivo}: {a.erro}" for a in resultados if a.erro]
    md = [
        "# Conferência automática gabarito x PDFs",
        "",
        f"- PDFs: {sum(not a.ausente and not a.erro for a in resultados)} lidos, "
        f"{len(nao_lidos)} não lidos",
        f"- Notas do gabarito conferidas: {total_notas} ({notas_ok} sem divergência)",
        f"- Verificações: {verif} — OK {verif_ok}, divergentes {len(divergentes)}, "
        f"não verificáveis {nao_verif}",
        "",
        "## Divergências",
        "",
        "| Arquivo | Nota | Data | Verificação | Gabarito | PDF | Obs. |",
        "|---|---|---|---|--:|--:|---|",
        *(divergentes or ["| — | | | | | | |"]),
        "",
        "## Páginas com negócios sem nota correspondente no gabarito",
        "",
        *(orfas or ["- nenhuma"]),
        "",
        "## PDFs não lidos",
        "",
        *(nao_lidos or ["- nenhum"]),
        "",
    ]
    return "\n".join(md), csv
