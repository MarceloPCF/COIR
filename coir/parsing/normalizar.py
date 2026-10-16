"""Nota extraída -> operações normalizadas (ticker, mercado, categoria, total).

Sem chute: nome de pregão sem cadastro levanta ReferenciaError com sugestões (rapidfuzz) para o
usuário cadastrar em data/ativos.csv. Exercício de opção: ativo-objeto = 4 letras do código + dígito
da classe (ON->3, PN->4, ...); sem classe na linha, a classe da abertura da série; por último
data/opcoes_raiz.csv.
"""

import re
from collections.abc import Iterable

from rapidfuzz import process

from coir.domain.nota import NegocioBMF, NegocioBovespa, NotaExtraida, OperacaoExtraida, Segmento
from coir.domain.referencias import Referencias
from coir.erros import ReferenciaError
from coir.parsing.numeros import compactar

RE_TICKER = re.compile(r"^[A-Z0-9]{4}\d{1,2}F?$")
RE_OPCAO = re.compile(
    r"^(?P<serie>[A-Z0-9]{4}[A-X]\d{1,4}(?:W\d)?)(?P<e>E)?(?P<classe>ON|PN[A-C]?|UNT)?$"
)  # W1..W5: opção semanal
TIPOS_BMF_IGNORADOS = (
    "AJUPOS",
    "PERMAN",
)  # ajuste de posição e taxa de permanência não são negócios


def classes_das_series(notas: Iterable[NotaExtraida], refs: Referencias) -> dict[str, str]:
    """Série de opção -> classe do ativo-objeto, vista na abertura (ex.: 'BBASN380 ON' -> ON)."""
    series: dict[str, str] = {}
    for nota in sorted(notas, key=lambda n: n.data_pregao):
        for n in nota.negocios_bovespa:
            tokens = n.especificacao.split()
            m = RE_OPCAO.match(tokens[0]) if tokens else None
            if m and len(tokens) > 1 and tokens[1] in refs.classes:
                series.setdefault(m["serie"], tokens[1])
    return series


def operacoes(
    nota: NotaExtraida, refs: Referencias, series: dict[str, str]
) -> list[OperacaoExtraida]:
    if nota.segmento is Segmento.BMF:
        return [_bmf(nota, n, refs) for n in nota.negocios_bmf if _negocio_bmf(n)]
    return [_bovespa(nota, n, refs, series) for n in nota.negocios_bovespa]


def _negocio_bmf(n: NegocioBMF) -> bool:
    return n.quantidade != 0 and not any(t in n.tipo.upper() for t in TIPOS_BMF_IGNORADOS)


def _comum(nota: NotaExtraida) -> dict[str, object]:
    return {"corretora": nota.corretora, "conta": nota.conta, "nota": nota.numero,
            "data": nota.data_pregao}  # fmt: skip


def _bmf(nota: NotaExtraida, n: NegocioBMF, refs: Referencias) -> OperacaoExtraida:
    raiz = n.mercadoria[:3]
    contrato = next(
        (c for c in refs.contratos if c.raiz == raiz and c.vigente(nota.data_pregao)), None
    )
    if contrato is None:
        raise ReferenciaError(
            f"contrato futuro {n.mercadoria} sem cadastro em contratos_futuros.csv"
        )
    return OperacaoExtraida(
        **_comum(nota),  # type: ignore[arg-type]
        categoria="DAY_TRADE" if n.day_trade else "COMUM",
        cv=n.cv,
        papel=n.mercadoria,
        mercado=contrato.mercado,
        preco=n.preco,
        quantidade=n.quantidade,
        total=n.preco * n.quantidade * contrato.multiplicador,
        ativo_objeto=None,
    )


def _bovespa(
    nota: NotaExtraida, n: NegocioBovespa, refs: Referencias, series: dict[str, str]
) -> OperacaoExtraida:
    mercado = refs.mercados_bovespa.get(n.mercado.upper())
    if mercado is None:
        raise ReferenciaError(f"tipo de mercado {n.mercado!r} sem cadastro em mercados_bovespa.csv")
    impresso, objeto = _papel(n, mercado, refs, series)
    papel = refs.ticker_na_data(impresso, nota.data_pregao)
    return OperacaoExtraida(
        **_comum(nota),  # type: ignore[arg-type]
        categoria="DAY_TRADE" if n.day_trade else "COMUM",
        cv=n.cv,
        papel=papel,
        mercado=mercado,
        preco=n.preco,
        quantidade=n.quantidade,
        total=n.valor,
        ativo_objeto=objeto,
        prazo_termo_dias=int(n.prazo) if mercado == "TERMO" and n.prazo.isdigit() else None,
        ticker_impresso=impresso if impresso != papel else None,
    )


def _papel(
    n: NegocioBovespa, mercado: str, refs: Referencias, series: dict[str, str]
) -> tuple[str, str | None]:
    tokens = n.especificacao.split()
    primeiro = tokens[0] if tokens else ""
    if mercado in ("CALL", "PUT", "EXERC CALL", "EXERC PUT"):
        m = RE_OPCAO.match(primeiro)
        if not m:
            raise ReferenciaError(f"código de opção inválido: {n.especificacao!r}")
        classe = m["classe"] or (tokens[1] if len(tokens) > 1 else None) or series.get(m["serie"])
        objeto = _ativo_objeto(m["serie"], classe, refs)
        if mercado.startswith("EXERC"):
            if objeto is None:
                raise ReferenciaError(f"exercício {primeiro}: ativo-objeto não identificado")
            return objeto, m["serie"]
        return m["serie"], objeto
    if RE_TICKER.match(primeiro):  # BTG imprime o código (às vezes o fracionário "LREN3F" na VISTA)
        return _sem_fracionario(primeiro), None
    return _ticker_por_nome(n.especificacao, refs), None


def _ativo_objeto(serie: str, classe: str | None, refs: Referencias) -> str | None:
    if classe and classe in refs.classes:
        return serie[:4] + refs.classes[classe]
    return refs.opcoes_raiz.get(serie[:4])


def _sem_fracionario(ticker: str) -> str:
    """Código do fracionário (final F) é o mesmo ativo do lote padrão: mesmo PM."""
    return ticker[:-1] if ticker.endswith("F") else ticker


def _ticker_por_nome(especificacao: str, refs: Referencias) -> str:
    chave = compactar(especificacao)
    while True:
        if chave in refs.ativos:
            return refs.ativos[chave]
        sufixo = next((s for s in refs.sufixos if chave.endswith(s) and len(chave) > len(s)), None)
        if sufixo is None:
            break
        chave = chave[: -len(sufixo)]
    sugestoes = [
        s for s, _, _ in process.extract(compactar(especificacao), list(refs.ativos), limit=3)
    ]
    raise ReferenciaError(
        f"ativo {especificacao!r} sem cadastro em data/ativos.csv "
        f"(parecidos: {', '.join(refs.ativos[s] + '=' + s for s in sugestoes)})"
    )
