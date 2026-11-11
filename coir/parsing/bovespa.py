"""Gramática Bovespa (SINACOR) com as variantes vistas nas notas: Ágora, grupo XP (Rico, XP antiga e
nova), BTG e "b3rv" (Clear, com negócios "B3 RV LISTADO" e o resumo em blocos). A linha de negócio e
o resumo seguem a mesma ordem; muda onde o valor fica em relação ao rótulo e onde estão cabeçalho e
conta.
"""

import re
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import date
from decimal import ROUND_DOWN, ROUND_HALF_UP, Decimal
from typing import Literal

from coir.domain.nota import (
    DC,
    CategoriaCusto,
    Custo,
    GrupoCusto,
    Montante,
    NegocioBovespa,
    ResumoBovespa,
)
from coir.erros import ExtracaoError
from coir.parsing import observacoes, rotulos
from coir.parsing.numeros import (
    RE_DATA,
    RE_DECIMAL,
    RE_INTEIRO,
    RE_MONTANTE,
    data_br,
    decimal_br,
    montantes,
)
from coir.parsing.rotulos import padrao, procurar

Variante = Literal["agora", "xp", "btg", "b3rv"]

MERCADOS = (
    r"OPCAO\s*DE\s*COMPRA|OPCAO\s*DE\s*VENDA|EXERC\s*OPC\s*COMPRA|EXERC\s*OPC\s*VENDA|"
    r"VISTA|FRACIONARIO|TERMO|LEILAO"
)
_FIM_XP = (
    rf"(?P<qtd>{RE_INTEIRO})\s+(?P<preco>{RE_DECIMAL})\s+(?P<valor>{RE_DECIMAL})\s+(?P<dc>[DC])"
)
_FIM_AGORA = (
    rf"(?P<qtd>{RE_INTEIRO})\s+(?P<preco>{RE_DECIMAL})\s+(?P<dc>[DC])\s+(?P<valor>{RE_DECIMAL})"
)
# o layout b3rv imprime a praça "B3 RV LISTADO" colada no C/V: "B3 RV LISTADOC VISTA"
_PRACA = r"(?:(?:\d-)?BOVESPA\s+|B3\s*RV\s*LISTADO\s*)"
_INICIO = rf"^\s*{_PRACA}(?P<cv>[CV])\s+(?P<mercado>{MERCADOS})\s+(?P<meio>.*?)\s+"
RE_INICIO_NEGOCIO = re.compile(rf"^\s*{_PRACA}[CV]\s")
RE_NEGOCIO = {
    "xp": re.compile(_INICIO + _FIM_XP + r"\s*$"),
    "btg": re.compile(_INICIO + _FIM_XP + r"\s*$"),
    "agora": re.compile(_INICIO + _FIM_AGORA + r"\s*$"),
    "b3rv": re.compile(_INICIO + _FIM_XP + r"\s*$"),
}
RE_PRAZO = re.compile(r"^(?P<prazo>\d{2}/\d{2}|\d{1,4})\s+")
RE_LIQUIDO = re.compile(
    rf"{padrao('liquido_para')}\s*(?P<data>\d{{2}}/\d{{2}}/\d{{4}})\s*"
    rf"(?P<valor>{RE_DECIMAL})\s*\|?\s*(?P<dc>[DC])?"
)
RE_IRRF = re.compile(
    rf"{padrao('irrf_operacoes')}\s*,?\s*base\s*(?:R\$)?\s*(?P<base>{RE_DECIMAL})"
    rf"(?:\s+(?P<valor>{RE_DECIMAL})(?:\s*(?P<dc>[DC]))?)?"
)
# a XP cola "IRRFDay Trade": aqui o rótulo pode começar no meio da palavra
_DAY_TRADE_PROJECAO = padrao("day_trade", colado=True), padrao("projecao")
RE_IRRF_DT = re.compile(
    rf"{_DAY_TRADE_PROJECAO[0]}.*?{_DAY_TRADE_PROJECAO[1]}\s*R\$\s*(?P<valor>{RE_DECIMAL})"
)
RE_TEM_IRRF_DT = re.compile(rf"{_DAY_TRADE_PROJECAO[0]}[\s\S]*?{_DAY_TRADE_PROJECAO[1]}")
RE_ISS_AGORA = re.compile(
    rf"{padrao('iss')}\s*\([^)]*\)\s*(?P<valor>{RE_DECIMAL})(?:\s*(?P<dc>[DC]))?"
)
RE_SO_VALOR = re.compile(rf"^\s*{RE_MONTANTE.pattern}\s*$")
RE_CONTA = re.compile(r"^\s*(?P<codigo>\d{1,4}-\d)\s+(?P<conta>\d+)\s*$")
RE_CONTA_AGORA = re.compile(r"^\s*(?P<conta>\d+)-(?P<dv>\d)\s+\D")
RE_CONTA_B3RV = re.compile(r"^\s*\d{1,4}\s*-\s*\d\s+(?P<conta>\d+)\s*-\s*\d(?:\s+\d+)?\s*$")
RE_LIQUIDO_B3RV = re.compile(
    rf"{padrao('liquido_para')}\s*(?P<valor>{RE_DECIMAL})\s*(?P<dc>[DC])\s*$"
)
# preço cotado por lote: valor = quantidade x preço / fator (RAIZ55: 100; lote de mil: "*")
FATORES_COTACAO = (1, 10, 100, 1000)
CENTAVO = Decimal("0.01")

# (rótulos em rotulos.yaml, nome gravado no custo, categoria, grupo); vale o primeiro rótulo que
# tiver valor na nota
RegraCusto = tuple[tuple[str, ...], str, CategoriaCusto, GrupoCusto]
CUSTOS_COMUNS: tuple[RegraCusto, ...] = (
    (("taxa_liquidacao",), "Taxa de liquidação/CCP", CategoriaCusto.LIQUIDACAO, "cblc"),
    (("taxa_registro",), "Taxa de Registro", CategoriaCusto.REGISTRO, "cblc"),
    (("taxa_termo_opcoes",), "Taxa de termo/opções", CategoriaCusto.TERMO_OPCOES, "bolsa"),
    (("taxa_ana",), "Taxa A.N.A.", CategoriaCusto.ANA, "bolsa"),
    (("emolumentos",), "Emolumentos", CategoriaCusto.EMOLUMENTOS, "bolsa"),
)
CUSTOS_DESPESAS: dict[Variante, tuple[RegraCusto, ...]] = {
    "agora": (
        (("corretagem",), "Corretagem", CategoriaCusto.CORRETAGEM, "despesas"),
        (("outras",), "Outras", CategoriaCusto.OUTROS, "despesas"),
    ),
    "xp": (
        (("taxa_transferencia_ativos",), "Taxa de Transf. de Ativos", CategoriaCusto.TRANSFERENCIA,
         "bolsa"),
        (("taxa_operacional", "corretagem"), "Taxa Operacional", CategoriaCusto.CORRETAGEM,
         "despesas"),
        (("execucao",), "Execução", CategoriaCusto.EXECUCAO, "despesas"),
        (("taxa_custodia",), "Taxa de Custódia", CategoriaCusto.CUSTODIA, "despesas"),
        (("impostos",), "Impostos", CategoriaCusto.IMPOSTOS, "despesas"),
        (("outros",), "Outros", CategoriaCusto.OUTROS, "despesas"),
    ),
    "btg": (
        (("taxa_transferencia_ativos",), "Taxa de Transferencia de Ativos",
         CategoriaCusto.TRANSFERENCIA, "cblc"),
        (("clearing",), "Clearing", CategoriaCusto.CORRETAGEM, "despesas"),
        (("execucao",), "Execução", CategoriaCusto.EXECUCAO, "despesas"),
        (("execucao_casa",), "Execução casa", CategoriaCusto.EXECUCAO, "despesas"),
    ),
}  # fmt: skip
# layout b3rv, por bloco do resumo financeiro: (rótulo, nome gravado, categoria); o começo da
# linha basta ("ISS ( SÃO PAULO )") e "Execução casa" vem antes de "Execução"
CUSTOS_B3RV: dict[GrupoCusto, tuple[tuple[str, str, CategoriaCusto], ...]] = {
    "cblc": (
        ("taxa_liquidacao", "Taxa de liquidação/CCP", CategoriaCusto.LIQUIDACAO),
        ("taxa_registro", "Taxa de Registro", CategoriaCusto.REGISTRO),
    ),
    "bolsa": (
        ("taxa_termo_opcoes", "Taxa de termo/opções", CategoriaCusto.TERMO_OPCOES),
        ("taxa_ana", "Taxa A.N.A.", CategoriaCusto.ANA),
        ("emolumentos", "Emolumentos", CategoriaCusto.EMOLUMENTOS),
    ),
    "despesas": (
        ("clearing", "Clearing", CategoriaCusto.CORRETAGEM),
        ("execucao_casa", "Execução casa", CategoriaCusto.EXECUCAO),
        ("execucao", "Execução", CategoriaCusto.EXECUCAO),
        ("iss", "ISS", CategoriaCusto.IMPOSTOS),
        ("outros", "Outros", CategoriaCusto.OUTROS),
    ),
}
# rótulos do resumo dos negócios, na ordem impressa (XP e b3rv leem os 8 valores pela posição)
RESUMO_NEGOCIOS = (
    "debentures",
    "vendas_a_vista",
    "compras_a_vista",
    "opcoes_compras",
    "opcoes_vendas",
    "operacoes_a_termo",
    "valor_operacoes_titulos_publicos",  # sem uso no cálculo: as outras gramáticas nem o procuram
    "valor_operacoes",
)


@dataclass(frozen=True)
class Cabecalho:
    numero: str
    folha: int
    data_pregao: date


# --------------------------------------------------------------------------- cabeçalho e conta
def _linhas(texto: str) -> list[str]:
    return [linha for linha in texto.splitlines() if linha.strip()]


def numero_nota(texto: str) -> str:
    """'38.724' / '057990' -> '38724' / '57990' (como o gabarito guarda)."""
    digitos = texto.strip().replace(".", "")
    if not digitos.isdigit():
        raise ExtracaoError(f"número de nota inválido: {texto!r}")
    return str(int(digitos))


def _folha(texto: str) -> int:
    if not texto.strip().isdigit():
        raise ExtracaoError(f"número de folha inválido: {texto.strip()!r}")
    return int(texto)


def _vizinha(linhas: Sequence[str], i: int, rotulo: str) -> str:
    if not 0 <= i < len(linhas):
        raise ExtracaoError(f"cabeçalho: valor de {rotulo!r} ausente")
    return linhas[i]


def cabecalho(texto: str, variante: Variante) -> Cabecalho:
    """Rótulo quebrado em duas linhas: o valor de depois vem após a segunda, o de antes, antes da
    primeira."""
    linhas = _linhas(texto)
    if variante == "b3rv":  # número antes dos rótulos, data depois, folha antes do título da nota
        ini, fim = procurar(linhas, "numero_nota", "data_pregao")
        i_titulo = procurar(linhas, "nota_de_corretagem", inicio=ini)[0]
        return Cabecalho(
            numero_nota(_vizinha(linhas, ini - 1, "Nr. nota")),
            _folha(_vizinha(linhas, i_titulo - 1, "Folha")),
            data_br(_vizinha(linhas, fim + 1, "Data pregão")),
        )
    if variante == "agora":
        _, fim = procurar(linhas, "numero_nota", "folha", "data_pregao")
        partes = _vizinha(linhas, fim + 1, "Nr.Nota").split()
        if len(partes) < 3:
            raise ExtracaoError("cabeçalho da nota (Ágora) incompleto")
        return Cabecalho(numero_nota(partes[0]), _folha(partes[1]), data_br(partes[2]))
    btg = variante == "btg"  # BTG: o valor vem antes do rótulo

    def vizinha(achado: tuple[int, int], rotulo: str) -> str:
        return _vizinha(linhas, achado[0] - 1 if btg else achado[1] + 1, rotulo)

    nota = procurar(linhas, "numero_nota")
    folha = procurar(linhas, "folha", inicio=0 if btg else nota[0])
    data = procurar(linhas, "data_pregao", inicio=0 if btg else folha[0])
    return Cabecalho(
        numero_nota(vizinha(nota, "Nr. nota")),
        _folha(vizinha(folha, "Folha")),
        data_br(vizinha(data, "Data pregão")),
    )


def conta(texto: str, variante: Variante) -> str:
    """Conta do cliente sem zeros à esquerda (como o gabarito guarda)."""
    linhas = _linhas(texto)
    if variante == "b3rv":
        m = RE_CONTA_B3RV.match(
            _vizinha(linhas, procurar(linhas, "codigo_cliente")[1] + 1, "conta")
        )
        if m is None:
            raise ExtracaoError("conta do cliente não encontrada")
        return str(int(m["conta"]))
    if variante == "agora":
        for linha in reversed(linhas):
            m = RE_CONTA_AGORA.match(linha)
            if m:
                return str(int(m["conta"]))  # sem o dígito, como no gabarito
        raise ExtracaoError("conta do cliente (Ágora) não encontrada")
    for linha in reversed(linhas):
        m = RE_CONTA.match(linha)
        if m:
            return str(int(m["conta"]))
    raise ExtracaoError("conta do cliente não encontrada")


# ---------------------------------------------------------------------------------- negócios
def negocios(texto: str, variante: Variante) -> list[NegocioBovespa]:
    saida: list[NegocioBovespa] = []
    marcas = observacoes.legenda(texto)
    for linha in texto.splitlines():
        if "BOVESPA" not in linha and "B3" not in linha:
            continue
        m = RE_NEGOCIO[variante].match(linha)
        if not m:
            if RE_INICIO_NEGOCIO.match(linha):
                raise ExtracaoError(f"linha de negócio fora do padrão: {linha.strip()!r}")
            continue
        saida.append(_negocio(m, marcas))
    return saida


def _confere_valor(
    quantidade: Decimal, preco: Decimal, valor: Decimal, lote_mil: bool, especificacao: str
) -> None:
    """O valor impresso prevalece: quantidade x preço só confere a leitura. Vale o valor exato
    (arredondado ou truncado ao centavo) e também a diferença que o arredondamento do preço
    explica, como no preço médio de execuções agregadas (3 a 17,45 impressos com valor 52,34). O
    fator é o do lote de cotação (RAIZ55 cota por 100; a Ágora marca o lote de mil com "*")."""
    fatores = (1000,) if lote_mil else FATORES_COTACAO
    for fator in fatores:
        bruto = quantidade * preco / fator
        if valor in {bruto.quantize(CENTAVO, ROUND_HALF_UP), bruto.quantize(CENTAVO, ROUND_DOWN)}:
            return
    expoente = preco.as_tuple().exponent
    casas = -expoente if isinstance(expoente, int) else 0
    meio_tick = Decimal(5) / 10 ** (casas + 1)
    # só o fator conhecido da linha: lote com o valor exato acima já passou, e fator solto com folga
    # esconderia preço ou quantidade lidos errado
    fator = fatores[0]
    if abs(quantidade * preco / fator - valor) <= quantidade * meio_tick / fator + CENTAVO:
        return
    raise ExtracaoError(
        f"negócio {especificacao}: quantidade x preço = {quantidade * preco / fatores[0]} difere "
        f"do valor {valor} além do que o arredondamento do preço explica"
    )


def _negocio(m: re.Match[str], marcas: frozenset[str] = frozenset()) -> NegocioBovespa:
    mercado = re.sub(r"\s+", " ", m["mercado"]).strip()
    meio = m["meio"].strip()
    prazo = ""
    p = RE_PRAZO.match(meio)
    a_vista = mercado in {"VISTA", "FRACIONARIO"}  # "01/00" neles é ruído (XP 2017, Clear 2018)
    if p and (not a_vista or "/" in p["prazo"]):
        prazo = "" if a_vista else p["prazo"]
        meio = meio[p.end() :]
    tokens, obs, lote_mil = observacoes.separar(meio.split(), marcas)
    especificacao = " ".join(tokens)
    quantidade, preco, valor = decimal_br(m["qtd"]), decimal_br(m["preco"]), decimal_br(m["valor"])
    _confere_valor(quantidade, preco, valor, lote_mil, especificacao)
    return NegocioBovespa(
        cv=m["cv"],  # type: ignore[arg-type]
        mercado=mercado,
        prazo=prazo,
        especificacao=especificacao,
        observacao=obs,
        quantidade=quantidade,
        preco=preco,
        valor=valor,
        dc=m["dc"],  # type: ignore[arg-type]
        lote_mil=lote_mil,
    )


# ---------------------------------------------------------------------------------- resumo
def _valor(texto: str, chaves: Sequence[str]) -> Montante | None:
    for chave in chaves:
        v = rotulos.valor_apos(texto, chave)
        if v is not None:
            return v
    return None


def _obrigatorio(texto: str, chave: str) -> Montante:
    v = rotulos.valor_apos(texto, chave)
    if v is None:
        raise ExtracaoError(f"valor de {rotulos.referencia(chave)!r} não encontrado no resumo")
    return v


def _bloco_antes(linhas: list[str], chave: str) -> list[Montante]:
    """Valores (um por linha) logo antes da linha do rótulo (XP: resumo dos negócios)."""
    i = procurar(linhas, chave)[0]
    valores: list[Montante] = []
    j = i - 1
    while j >= 0 and RE_SO_VALOR.match(linhas[j]):
        valores.insert(0, montantes(linhas[j])[0])
        j -= 1
    return valores


def _bloco_depois(linhas: list[str], chave: str) -> list[Montante]:
    """Valores (um por linha) logo depois da linha do rótulo (BTG: ISS, IRRF e Outras)."""
    alvo = rotulos.regex(chave)
    for i, linha in enumerate(linhas):
        if alvo.search(linha):
            valores: list[Montante] = []
            for seguinte in linhas[i + 1 :]:
                if not RE_SO_VALOR.match(seguinte):
                    break
                valores.append(montantes(seguinte)[0])
            return valores
    return []


def _custo(rotulo: str, categoria: CategoriaCusto, grupo: GrupoCusto, v: Montante) -> Custo:
    return Custo(
        rotulo=rotulo, categoria=categoria, grupo=grupo, valor=-v.assinado if v.dc else v.valor
    )


def _colunas(trecho: Sequence[str]) -> list[tuple[str, Montante]]:
    """Bloco do layout b3rv: os rótulos, depois os valores, depois os D/C, cada um em sua linha, na
    mesma ordem. O valor pode vir na linha do rótulo; o IRRF, com base e valor, é lido à parte. Sem
    um valor e um D/C por rótulo, a nota não entra: casar errado trocaria um custo por outro."""
    rotulos: list[tuple[str, str | None]] = []
    soltos: list[str] = []
    indicadores: list[DC] = []
    for linha in trecho:
        s = linha.strip()
        if s == "D" or s == "C":
            indicadores.append(s)
        elif RE_SO_VALOR.match(s):
            soltos.append(s)
        elif not RE_IRRF.search(s):
            na_linha = re.search(rf"\s(?P<valor>{RE_DECIMAL})$", s)
            rotulos.append((s[: na_linha.start()], na_linha["valor"]) if na_linha else (s, None))
    faltam_valor = sum(valor is None for _, valor in rotulos)
    if faltam_valor != len(soltos) or len(rotulos) != len(indicadores):
        raise ExtracaoError(
            f"resumo financeiro: {len(rotulos)} rótulos para {len(soltos)} valores soltos e "
            f"{len(indicadores)} D/C"
        )
    fila = iter(soltos)
    return [
        (rotulo, Montante(valor=decimal_br(valor or next(fila)), dc=dc))
        for (rotulo, valor), dc in zip(rotulos, indicadores, strict=True)
    ]


def _bloco_custos(
    colunas: Sequence[tuple[str, Montante]], grupo: GrupoCusto
) -> tuple[list[Custo], Montante]:
    """Custos e total de um bloco; rótulo desconhecido recusa a nota, para não perder um custo."""
    custos: list[Custo] = []
    total: Montante | None = None
    for rotulo, valor in colunas:
        if rotulos.linha_inteira(rotulo, "total"):
            total = valor
            continue
        for chave, nome, categoria in CUSTOS_B3RV[grupo]:
            if rotulos.no_inicio(rotulo, chave):
                custos.append(_custo(nome, categoria, grupo, valor))
                break
        else:
            raise ExtracaoError(f"resumo financeiro: custo desconhecido {rotulo!r}")
    if total is None:
        raise ExtracaoError(f"resumo financeiro: bloco {grupo} sem total")
    return custos, total


def _resumo_b3rv(texto: str) -> ResumoBovespa:
    """Layout b3rv. Os rótulos do resumo dos negócios saem fora da ordem no texto, mas os 8 valores
    seguem a ordem de sempre (RESUMO_NEGOCIOS), e a conciliação confere compras e vendas."""
    linhas = _linhas(texto)
    fin = procurar(linhas, "resumo_financeiro")
    # "Clearing" aqui é só o título do bloco
    i_cblc = procurar(linhas, "clearing", inicio=fin[0])[1] + 1
    bolsa = procurar(linhas, "bolsa", inicio=i_cblc)
    desp = procurar(linhas, "corretagem_despesas", inicio=bolsa[0])
    quadro = procurar(linhas, "resumo_negocios", inicio=desp[0])
    i_obs = procurar(linhas, "observacoes", inicio=quadro[0])[0]
    primeiro = _colunas(linhas[i_cblc : bolsa[0]])
    if not primeiro or not rotulos.no_inicio(primeiro[0][0], "valor_liquido_operacoes"):
        raise ExtracaoError("'Valor líquido das operações' não encontrado no resumo")
    custos_cblc, total_cblc = _bloco_custos(primeiro[1:], "cblc")
    custos_bolsa, total_bolsa = _bloco_custos(_colunas(linhas[bolsa[1] + 1 : desp[0]]), "bolsa")
    custos_desp, total_desp = _bloco_custos(_colunas(linhas[desp[1] + 1 : quadro[0]]), "despesas")
    custos = [*custos_cblc, *custos_bolsa, *custos_desp]
    transferencia = rotulos.valor_apos(texto, "taxa_transferencia_ativos")
    if transferencia is not None:
        custos.append(
            _custo("Taxa de Transferência de Ativos", CategoriaCusto.TRANSFERENCIA, "cblc",
                   transferencia)
        )  # fmt: skip
    valores = [montantes(s)[0].valor for s in linhas[quadro[1] + 1 : i_obs] if RE_SO_VALOR.match(s)]
    if len(valores) != len(RESUMO_NEGOCIOS):
        raise ExtracaoError(f"resumo dos negócios com {len(valores)} valores (esperado 8)")
    negs = dict(zip(RESUMO_NEGOCIOS, valores, strict=True))
    i_liq = next((i for i, s in enumerate(linhas) if RE_LIQUIDO_B3RV.search(s)), None)
    data = next((s for s in linhas[(i_liq or 0) + 1 :] if RE_DATA.fullmatch(s.strip())), None)
    if i_liq is None or data is None:
        raise ExtracaoError("'Líquido para' não encontrado")
    liquido = RE_LIQUIDO_B3RV.search(linhas[i_liq])
    assert liquido is not None
    irrf, base = _irrf(texto, "b3rv")
    return ResumoBovespa(
        valor_operacoes=negs["valor_operacoes"],
        compras=negs["compras_a_vista"] + negs["opcoes_compras"],
        vendas=negs["vendas_a_vista"] + negs["opcoes_vendas"],
        termo=negs["operacoes_a_termo"],
        valor_liquido_operacoes=primeiro[0][1],
        custos=tuple(custos),
        total_cblc=total_cblc,
        total_bovespa=total_bolsa,
        total_despesas=total_desp,
        irrf=irrf,
        irrf_base=base,
        irrf_no_liquido=False,  # decidido na conciliação
        irrf_day_trade=_irrf_day_trade(texto),
        liquido=Montante(valor=decimal_br(liquido["valor"]), dc=liquido["dc"]),  # type: ignore[arg-type]
        data_liquidacao=data_br(data),
    )


def _irrf(texto: str, variante: Variante) -> tuple[Decimal, Decimal | None]:
    """Valor e base do IRRF sobre operações, na linha "I.R.R.F. s/ operações, base ..."."""
    m_irrf = RE_IRRF.search(texto)
    if m_irrf is None:
        if rotulos.regex("irrf_operacoes", colado=True).search(texto):
            raise ExtracaoError("linha de I.R.R.F. s/ operações fora do padrão")
        return Decimal(0), None
    if m_irrf["valor"]:
        return decimal_br(m_irrf["valor"]), decimal_br(m_irrf["base"])
    if variante != "btg":  # só a BTG imprime o valor do IRRF longe da base
        raise ExtracaoError("I.R.R.F. s/ operações sem valor legível após a base")
    return Decimal(0), decimal_br(m_irrf["base"])


def _irrf_day_trade(texto: str) -> Decimal:
    m_dt = RE_IRRF_DT.search(texto)
    if m_dt is None and RE_TEM_IRRF_DT.search(texto):
        raise ExtracaoError("linha de IRRF day trade (projeção) fora do padrão")
    return decimal_br(m_dt["valor"]) if m_dt else Decimal(0)


def resumo(texto: str, variante: Variante) -> ResumoBovespa:
    """Resumo dos negócios + resumo financeiro da última folha da nota."""
    if variante == "b3rv":
        return _resumo_b3rv(texto)
    linhas = _linhas(texto)
    if variante == "xp":
        bloco = _bloco_antes(linhas, "resumo_negocios")
        if len(bloco) != len(RESUMO_NEGOCIOS):
            raise ExtracaoError(f"resumo dos negócios com {len(bloco)} valores (esperado 8)")
        negs = dict(zip(RESUMO_NEGOCIOS, (b.valor for b in bloco), strict=True))
    else:
        negs = {c: _obrigatorio(texto, c).valor for c in RESUMO_NEGOCIOS
                if c != "valor_operacoes_titulos_publicos"}  # fmt: skip
    custos: list[Custo] = []
    for chaves, nome, categoria, grupo in (*CUSTOS_COMUNS, *CUSTOS_DESPESAS[variante]):
        v = _valor(texto, chaves)
        if v is not None:
            custos.append(_custo(nome, categoria, grupo, v))
    irrf, base = _irrf(texto, variante)
    if variante == "agora":
        iss = RE_ISS_AGORA.search(texto)
        if iss:
            custos.append(
                Custo(
                    rotulo="ISS",
                    categoria=CategoriaCusto.IMPOSTOS,
                    grupo="despesas",
                    valor=decimal_br(iss["valor"]),
                )
            )
    if variante == "btg":  # ISS, IRRF e Outras: valores sem rótulo depois de "Execução casa"
        orfaos = _bloco_depois(linhas, "execucao_casa")
        if len(orfaos) != 3:
            raise ExtracaoError(
                f"BTG: {len(orfaos)} valores sem rótulo após Execução casa (esperado 3)"
            )
        custos.append(_custo("ISS", CategoriaCusto.IMPOSTOS, "despesas", orfaos[0]))
        irrf = orfaos[1].valor
        custos.append(_custo("Outras Bovespa", CategoriaCusto.OUTROS, "despesas", orfaos[2]))
    liquido = RE_LIQUIDO.search(texto)
    if not liquido:
        raise ExtracaoError("'Líquido para' não encontrado")
    vlo = _obrigatorio(texto, "valor_liquido_operacoes")
    return ResumoBovespa(
        valor_operacoes=negs["valor_operacoes"],
        compras=negs["compras_a_vista"] + negs["opcoes_compras"],
        vendas=negs["vendas_a_vista"] + negs["opcoes_vendas"],
        termo=negs["operacoes_a_termo"],
        valor_liquido_operacoes=vlo,
        custos=tuple(custos),
        total_cblc=_obrigatorio(texto, "total_cblc"),
        total_bovespa=_obrigatorio(texto, "total_bovespa"),
        total_despesas=_obrigatorio(texto, "total_despesas"),
        irrf=irrf,
        irrf_base=base,
        irrf_no_liquido=False,  # decidido na conciliação
        irrf_day_trade=_irrf_day_trade(texto),
        liquido=Montante(valor=decimal_br(liquido["valor"]), dc=liquido["dc"]),  # type: ignore[arg-type]
        data_liquidacao=data_br(liquido["data"]),
    )
