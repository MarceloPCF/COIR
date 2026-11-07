"""Nota extraída -> operações normalizadas (ticker, mercado, categoria, total).

Sem chute: nome de pregão que nenhuma tabela conhece levanta ReferenciaError dizendo o que fazer
(atualizar as tabelas; persistindo, mandar a nota pelo Suporte). O nome é procurado nas tabelas
curadas e no histórico de nomes de pregão da B3, pela data do negócio. Exercício de opção:
ativo-objeto = 4 letras do código + dígito
da classe (ON->3, PN->4, ...); sem classe na linha, a classe da abertura da série; por último
data/opcoes_raiz.csv.
"""

import re
from collections.abc import Iterable
from datetime import date

from rapidfuzz import process

from coir.domain.fiscal import TipoAjuste
from coir.domain.nota import NegocioBMF, NegocioBovespa, NotaExtraida, OperacaoExtraida, Segmento
from coir.domain.referencias import Referencias
from coir.erros import ReferenciaError
from coir.parsing import futuros
from coir.parsing.numeros import compactar
from coir.parsing.observacoes import marca_suspeita

RE_TICKER = re.compile(r"^[A-Z0-9]{4}\d{1,2}F?$")
RE_OPCAO = re.compile(
    r"^(?P<serie>[A-Z0-9]{4}[A-X]\d{1,4}(?:W\d)?)(?P<e>E)?(?P<classe>ON|PN[A-C]?|UNT)?$"
)  # W1..W5: opção semanal


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
        futuros.conversao_do_volume(nota, refs)  # confere o câmbio implícito de cada linha
        return [_bmf(nota, n, refs) for n in nota.negocios_bmf if _negocio_bmf(n)]
    return [_bovespa(nota, n, refs, series) for n in nota.negocios_bovespa]


def _negocio_bmf(n: NegocioBMF) -> bool:
    """Ajuste de posição e taxa de permanência não são negócios."""
    return n.quantidade != 0 and futuros.tipo(n) in (TipoAjuste.NORMAL, TipoAjuste.DAY_TRADE)


def _comum(nota: NotaExtraida) -> dict[str, object]:
    return {"corretora": nota.corretora, "conta": nota.conta, "nota": nota.numero,
            "data": nota.data_pregao}  # fmt: skip


def _bmf(nota: NotaExtraida, n: NegocioBMF, refs: Referencias) -> OperacaoExtraida:
    contrato = futuros.contrato(n, refs, nota.data_pregao)
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
    impresso, objeto = _papel(n, mercado, refs, series, nota.data_pregao)
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
    n: NegocioBovespa, mercado: str, refs: Referencias, series: dict[str, str], dia: date
) -> tuple[str, str | None]:
    tokens = n.especificacao.split()
    primeiro = tokens[0] if tokens else ""
    if mercado in ("CALL", "PUT", "EXERC CALL", "EXERC PUT"):
        m = RE_OPCAO.match(primeiro)
        if not m:
            raise ReferenciaError(f"código de opção inválido: {n.especificacao!r}")
        classe = m["classe"] or (tokens[1] if len(tokens) > 1 else None) or series.get(m["serie"])
        objeto = _ativo_objeto(m["serie"], classe, refs, dia)
        if mercado.startswith("EXERC"):
            if objeto is None:
                raise ReferenciaError(f"exercício {primeiro}: ativo-objeto não identificado")
            return objeto, m["serie"]
        return m["serie"], objeto
    if RE_TICKER.match(primeiro):  # BTG imprime o código (às vezes o fracionário "LREN3F" na VISTA)
        suspeita = next((t for t in tokens[1:] if marca_suspeita(t)), None)
        if suspeita is not None:
            raise ReferenciaError(_marca_de_day_trade(n.especificacao, suspeita))
        return _sem_fracionario(primeiro), None
    return _ticker_por_nome(n, refs, dia), None


def _ativo_objeto(serie: str, classe: str | None, refs: Referencias, dia: date) -> str | None:
    if classe and classe in refs.classes:
        objeto = serie[:4] + refs.classes[classe]
        # série de raiz que continua depois da troca da classe (RAIZ PN -> RAIZ55): o código novo
        if any(t.ticker_impresso == objeto and dia > t.ate for t in refs.trocas):
            return refs.ticker_na_data(objeto, date.min)
        return objeto
    return refs.opcoes_raiz.get(serie[:4])


def _sem_fracionario(ticker: str) -> str:
    """Código do fracionário (final F) é o mesmo ativo do lote padrão: mesmo PM."""
    return ticker[:-1] if ticker.endswith("F") else ticker


def _procurar(especificacao: str, refs: Referencias, dia: date) -> str | None:
    """Nome de pregão -> ticker, tirando os sufixos de governança do fim um a um: primeiro as
    tabelas curadas, depois o histórico da B3 pela data do negócio."""
    chave = compactar(especificacao)
    while True:
        if chave in refs.ativos:
            return refs.ativos[chave]
        ticker = refs.nome_na_data(chave, dia)
        if ticker is not None:
            return ticker
        sufixo = next((s for s in refs.sufixos if chave.endswith(s) and len(chave) > len(s)), None)
        if sufixo is None:
            return None
        chave = chave[: -len(sufixo)]


def _marca_de_day_trade(especificacao: str, marca: str) -> str:
    return (f"{especificacao!r}: {marca!r} parece marca de day trade que o COIR não conhece. Mande "
            "a nota pela tela Suporte; a nota não entrou para o day trade não virar operação "
            "comum")  # fmt: skip


def _ticker_por_nome(n: NegocioBovespa, refs: Referencias, dia: date) -> str:
    achado = _procurar(n.especificacao, refs, dia)
    # o começo das marcas pode ser o fim do nome: "RAIZEN PNA P", "GOOGLE DRN A" + D
    for k in range(len(n.observacao), 0, -1):
        if achado is None and n.observacao[:k].isalnum():
            achado = _procurar(f"{n.especificacao} {n.observacao[:k]}", refs, dia)
    tokens = n.especificacao.split()
    if achado is None and len(tokens) > 1 and len(tokens[-1]) <= 3:
        if marca_suspeita(tokens[-1]):
            raise ReferenciaError(_marca_de_day_trade(n.especificacao, tokens[-1]))
        achado = _procurar(" ".join(tokens[:-1]), refs, dia)  # marca nova, fora da legenda
    if achado is None:
        raise ReferenciaError(_sem_cadastro(n.especificacao, refs))
    return achado


def _sem_cadastro(especificacao: str, refs: Referencias) -> str:
    """Só sugere nome muito parecido: sugestão fraca já levou a procurar o ativo errado."""
    parecidos = process.extract(compactar(especificacao), list(refs.ativos), limit=3,
                                score_cutoff=90)  # fmt: skip
    nomes = ", ".join(refs.ativos[s] + "=" + s for s, _, _ in parecidos)
    dica = f" Parecidos, não usados: {nomes}."
    return (
        f"ativo {especificacao!r} não encontrado nas tabelas do COIR. Atualize as tabelas em "
        "Configurações e importe de novo; se continuar, mande a nota pela tela Suporte."
        + (dica if parecidos else "")
    )
