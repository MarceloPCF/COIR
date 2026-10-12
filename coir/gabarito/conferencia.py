"""Conferência automática gabarito x PDFs das notas, em todas as notas de todos os PDFs.

Para cada nota do gabarito, localiza as páginas do PDF com o número da nota e confere:
- Bovespa: por C/V, soma das quantidades e dos valores das linhas de negócio; líquido da nota
  ("Líquido para") contra vendas - compras - custos (com ou sem o IRRF descontado);
- BM&F: por C/V, soma das quantidades e de quantidade x preço; em notas só de day trade,
  ajuste day trade, IRRF e líquido esperados presentes entre os valores impressos na página.
Também lista páginas com negócios que não pertencem a nenhuma nota do gabarito.

É uma ferramenta de validação do gabarito (F1), não o extrator da F2: lê só o necessário.
"""

import re
from collections import defaultdict
from collections.abc import Iterable
from dataclasses import dataclass, field
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
from coir.parsing.texto_pdf import paginas

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
    erro: str = ""


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
    so_day_trade = all(qtd[(p, "C")] == qtd[(p, "V")] for p, _ in qtd)
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
        res.verificacoes.append(
            Verificacao(nome, valor, None, _presente(valor, impressos), "procurado no resumo")
        )
    return res


def _ops_por_nota(gabarito: Gabarito, arq: ArquivoNota) -> dict[str, list[OperacaoGabarito]]:
    por_nota: dict[str, list[OperacaoGabarito]] = defaultdict(list)
    for op in gabarito.operacoes:
        if (op.corretora, op.data.year, op.data.month) != (arq.corretora, arq.ano, arq.mes):
            continue
        if (op.mercado in MERCADOS_BMF) != arq.bmf:
            continue
        por_nota[op.nota].append(op)
    return dict(sorted(por_nota.items(), key=lambda kv: (kv[1][0].data, kv[0])))


def conferir_arquivo(gabarito: Gabarito, caminho: Path) -> ResultadoArquivo:
    res = ResultadoArquivo(caminho.name)
    arq = interpretar_nome(caminho.name)
    textos = paginas(caminho)
    usadas: set[int] = set()
    for nota, ops in _ops_por_nota(gabarito, arq).items():
        pags = paginas_da_nota(textos, nota)
        if not pags:
            r = ResultadoNota(caminho.name, nota, ops[0].data.isoformat(), ops[0].conta)
            r.verificacoes.append(Verificacao("nota no PDF", None, None, False, "não encontrada"))
            res.notas.append(r)
            continue
        usadas.update(pags)
        texto = "\n".join(textos[i] for i in pags)
        res.notas.append(conferir_nota(caminho.name, ops, texto, arq.bmf))
    # procura os dois tipos de negócio: há PDFs mistos (ações + BM&F no mesmo arquivo)
    res.paginas_sem_nota = [
        i + 1
        for i, t in enumerate(textos)
        if i not in usadas and (linhas_bovespa(t) or linhas_bmf(t))
    ]
    return res


def conferir_pasta(gabarito: Gabarito, pasta: Path) -> list[ResultadoArquivo]:
    resultados: list[ResultadoArquivo] = []
    for caminho in sorted(pasta.glob("*.pdf")):
        try:
            interpretar_nome(caminho.name)
        except GabaritoError:
            continue  # ex.: opção flexível, fora do escopo
        resultados.append(conferir_arquivo(gabarito, caminho))
    return resultados


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
    md = [
        "# Conferência automática gabarito x PDFs",
        "",
        f"- PDFs: {len(resultados)}",
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
    ]
    return "\n".join(md), csv
