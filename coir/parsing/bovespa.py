"""Gramática Bovespa (SINACOR) com as três variantes vistas nas notas: Ágora, grupo XP (Rico, XP
antiga e nova) e BTG. A linha de negócio e o resumo seguem a mesma ordem; muda onde o valor fica em
relação ao rótulo e onde estão cabeçalho e conta.
"""

import re
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import date
from decimal import ROUND_DOWN, ROUND_HALF_UP, Decimal
from typing import Literal

from coir.domain.nota import (
    CategoriaCusto,
    Custo,
    GrupoCusto,
    Montante,
    NegocioBovespa,
    ResumoBovespa,
)
from coir.erros import ExtracaoError
from coir.parsing.numeros import (
    RE_DECIMAL,
    RE_INTEIRO,
    RE_MONTANTE,
    compactar,
    data_br,
    decimal_br,
    montantes,
    regex_rotulo,
    valor_apos_rotulo,
)

Variante = Literal["agora", "xp", "btg"]

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
_INICIO = rf"^\s*(?:\d-)?BOVESPA\s+(?P<cv>[CV])\s+(?P<mercado>{MERCADOS})\s+(?P<meio>.*?)\s+"
RE_NEGOCIO = {
    "xp": re.compile(_INICIO + _FIM_XP + r"\s*$"),
    "btg": re.compile(_INICIO + _FIM_XP + r"\s*$"),
    "agora": re.compile(_INICIO + _FIM_AGORA + r"\s*$"),
}
RE_PRAZO = re.compile(r"^(?P<prazo>\d{2}/\d{2}|\d{1,4})\s+")
# marcas de observação (*) vistas nas notas: H home broker, @ contraparte, D day trade, # negócio
# direto, 2 corretora na contraparte, 8 liquidação institucional (ex.: "H#", "D#2", "@#")
RE_OBS = re.compile(r"^(?=.)H?@?D?#?[28]?$")
RE_LIQUIDO = re.compile(
    rf"L\s*[íi]\s*q\s*u\s*i\s*d\s*o\s*p\s*a\s*r\s*a\s*(?P<data>\d{{2}}/\d{{2}}/\d{{4}})\s*"
    rf"(?P<valor>{RE_DECIMAL})\s*\|?\s*(?P<dc>[DC])?"
)
RE_IRRF = re.compile(
    rf"I\.?\s*R\.?\s*R\.?\s*F\.?\s*s/\s*opera\S*\s*,?\s*base\s*R\$\s*(?P<base>{RE_DECIMAL})"
    rf"(?:\s+(?P<valor>{RE_DECIMAL})(?:\s*(?P<dc>[DC]))?)?"
)
RE_IRRF_DT = re.compile(
    rf"Day\s*-?\s*Trade.*?Proje\S*\s*R\$\s*(?P<valor>{RE_DECIMAL})", re.IGNORECASE
)
RE_ISS_AGORA = re.compile(rf"ISS\s*\([^)]*\)\s*(?P<valor>{RE_DECIMAL})(?:\s*(?P<dc>[DC]))?")
RE_SO_VALOR = re.compile(rf"^\s*{RE_MONTANTE.pattern}\s*$")
RE_CONTA = re.compile(r"^\s*(?P<codigo>\d{1,4}-\d)\s+(?P<conta>\d+)\s*$")
RE_CONTA_AGORA = re.compile(r"^\s*(?P<conta>\d+)-(?P<dv>\d)\s+\D")

# (rótulos aceitos, categoria, grupo); o primeiro rótulo que tiver valor na nota vale
CUSTOS_COMUNS: tuple[tuple[tuple[str, ...], CategoriaCusto, GrupoCusto], ...] = (
    (("Taxa de liquidação/CCP", "Taxa de liquidação"), CategoriaCusto.LIQUIDACAO, "cblc"),
    (("Taxa de Registro",), CategoriaCusto.REGISTRO, "cblc"),
    (("Taxa de termo/opções",), CategoriaCusto.TERMO_OPCOES, "bolsa"),
    (("Taxa A.N.A.",), CategoriaCusto.ANA, "bolsa"),
    (("Emolumentos",), CategoriaCusto.EMOLUMENTOS, "bolsa"),
)
CUSTOS_DESPESAS: dict[Variante, tuple[tuple[tuple[str, ...], CategoriaCusto, GrupoCusto], ...]] = {
    "agora": (
        (("Corretagem",), CategoriaCusto.CORRETAGEM, "despesas"),
        (("Outras",), CategoriaCusto.OUTROS, "despesas"),
    ),
    "xp": (
        (("Taxa de Transf. de Ativos",), CategoriaCusto.TRANSFERENCIA, "bolsa"),
        (("Taxa Operacional", "Corretagem"), CategoriaCusto.CORRETAGEM, "despesas"),
        (("Execução",), CategoriaCusto.EXECUCAO, "despesas"),
        (("Taxa de Custódia",), CategoriaCusto.CUSTODIA, "despesas"),
        (("Impostos",), CategoriaCusto.IMPOSTOS, "despesas"),
        (("Outros",), CategoriaCusto.OUTROS, "despesas"),
    ),
    "btg": (
        (("Taxa de Transferencia de Ativos",), CategoriaCusto.TRANSFERENCIA, "cblc"),
        (("Clearing",), CategoriaCusto.CORRETAGEM, "despesas"),
        (("Execução",), CategoriaCusto.EXECUCAO, "despesas"),
        (("Execução casa",), CategoriaCusto.EXECUCAO, "despesas"),
    ),
}
RESUMO_NEGOCIOS = (
    "Debêntures",
    "Vendas à vista",
    "Compras à vista",
    "Opções - compras",
    "Opções - vendas",
    "Operações à termo",
    "Valor das oper. c/ títulos públ.",
    "Valor das operações",
)


@dataclass(frozen=True)
class Cabecalho:
    numero: str
    folha: int
    data_pregao: date


# --------------------------------------------------------------------------- cabeçalho e conta
def _linhas(texto: str) -> list[str]:
    return [linha for linha in texto.splitlines() if linha.strip()]


def _indice(linhas: Sequence[str], rotulo: str, inicio: int = 0) -> int:
    alvo = compactar(rotulo)
    for i in range(inicio, len(linhas)):
        if alvo in compactar(linhas[i]):
            return i
    raise ExtracaoError(f"rótulo {rotulo!r} não encontrado")


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
    linhas = _linhas(texto)
    if variante == "agora":
        i = _indice(linhas, "Nr.Nota Folha Data pregão")
        partes = _vizinha(linhas, i + 1, "Nr.Nota").split()
        if len(partes) < 3:
            raise ExtracaoError("cabeçalho da nota (Ágora) incompleto")
        return Cabecalho(numero_nota(partes[0]), _folha(partes[1]), data_br(partes[2]))
    passo = -1 if variante == "btg" else 1  # BTG: o valor vem antes do rótulo
    i_nota = _indice(linhas, "Nr. nota")
    i_folha = _indice(linhas, "Folha", 0 if variante == "btg" else i_nota)
    i_data = _indice(linhas, "Data pregão", 0 if variante == "btg" else i_folha)
    return Cabecalho(
        numero_nota(_vizinha(linhas, i_nota + passo, "Nr. nota")),
        _folha(_vizinha(linhas, i_folha + passo, "Folha")),
        data_br(_vizinha(linhas, i_data + passo, "Data pregão")),
    )


def conta(texto: str, variante: Variante) -> str:
    """Conta do cliente sem zeros à esquerda (como o gabarito guarda)."""
    linhas = _linhas(texto)
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
    for linha in texto.splitlines():
        if "BOVESPA" not in linha:
            continue
        m = RE_NEGOCIO[variante].match(linha)
        if not m:
            if re.match(r"^\s*(?:\d-)?BOVESPA\s+[CV]\s", linha):
                raise ExtracaoError(f"linha de negócio fora do padrão: {linha.strip()!r}")
            continue
        saida.append(_negocio(m))
    return saida


def _negocio(m: re.Match[str]) -> NegocioBovespa:
    mercado = re.sub(r"\s+", " ", m["mercado"]).strip()
    meio = m["meio"].strip()
    prazo = ""
    p = RE_PRAZO.match(meio)
    if p and (mercado != "VISTA" or "/" in p["prazo"]):  # VISTA "01/00" (XP 2017) é ruído
        prazo = p["prazo"] if mercado != "VISTA" else ""
        meio = meio[p.end() :]
    tokens = meio.split()
    obs = ""
    if len(tokens) > 1 and RE_OBS.match(tokens[-1]):
        obs = tokens.pop()
    lote_mil = "*" in tokens
    especificacao = " ".join(t for t in tokens if t != "*")
    quantidade, preco, valor = decimal_br(m["qtd"]), decimal_br(m["preco"]), decimal_br(m["valor"])
    bruto = quantidade * preco / (1000 if lote_mil else 1)
    # valor impresso ao centavo: arredondado ou truncado (a Ágora trunca no lote de mil)
    centavo = Decimal("0.01")
    if valor not in {bruto.quantize(centavo, ROUND_HALF_UP), bruto.quantize(centavo, ROUND_DOWN)}:
        raise ExtracaoError(
            f"negócio {especificacao}: quantidade x preço = {bruto} difere do valor {valor}"
        )
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
def _valor(texto: str, rotulos: Sequence[str]) -> Montante | None:
    for rotulo in rotulos:
        v = valor_apos_rotulo(texto, rotulo)
        if v is not None:
            return v
    return None


def _obrigatorio(texto: str, *rotulos: str) -> Montante:
    v = _valor(texto, rotulos)
    if v is None:
        raise ExtracaoError(f"valor de {rotulos[0]!r} não encontrado no resumo")
    return v


def _bloco_antes(linhas: list[str], rotulo: str) -> list[Montante]:
    """Valores (um por linha) logo antes da linha do rótulo (XP: resumo dos negócios)."""
    i = _indice(linhas, rotulo)
    valores: list[Montante] = []
    j = i - 1
    while j >= 0 and RE_SO_VALOR.match(linhas[j]):
        valores.insert(0, montantes(linhas[j])[0])
        j -= 1
    return valores


def _bloco_depois(linhas: list[str], rotulo: str) -> list[Montante]:
    """Valores (um por linha) logo depois da linha do rótulo (BTG: ISS, IRRF e Outras)."""
    padrao = regex_rotulo(rotulo)
    for i, linha in enumerate(linhas):
        if padrao.search(linha):
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


def resumo(texto: str, variante: Variante) -> ResumoBovespa:
    """Resumo dos negócios + resumo financeiro da última folha da nota."""
    linhas = _linhas(texto)
    if variante == "xp":
        bloco = _bloco_antes(linhas, "Resumo dos Negócios")
        if len(bloco) != len(RESUMO_NEGOCIOS):
            raise ExtracaoError(f"resumo dos negócios com {len(bloco)} valores (esperado 8)")
        negs = dict(zip(RESUMO_NEGOCIOS, (b.valor for b in bloco), strict=True))
    else:
        negs = {r: _obrigatorio(texto, r).valor for r in RESUMO_NEGOCIOS if r != RESUMO_NEGOCIOS[6]}
    custos: list[Custo] = []
    for rotulos, categoria, grupo in (*CUSTOS_COMUNS, *CUSTOS_DESPESAS[variante]):
        v = _valor(texto, rotulos)
        if v is not None:
            custos.append(_custo(rotulos[0], categoria, grupo, v))
    irrf, base = Decimal(0), None
    m_irrf = RE_IRRF.search(texto)
    if m_irrf is None and "i.r.r.f.s/opera" in compactar(texto):
        raise ExtracaoError("linha de I.R.R.F. s/ operações fora do padrão")
    if m_irrf:
        base = decimal_br(m_irrf["base"])
        if m_irrf["valor"]:
            irrf = decimal_br(m_irrf["valor"])
        elif variante != "btg":  # só a BTG imprime o valor do IRRF longe da base
            raise ExtracaoError("I.R.R.F. s/ operações sem valor legível após a base")
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
        orfaos = _bloco_depois(linhas, "Execução casa")
        if len(orfaos) != 3:
            raise ExtracaoError(
                f"BTG: {len(orfaos)} valores sem rótulo após Execução casa (esperado 3)"
            )
        custos.append(_custo("ISS", CategoriaCusto.IMPOSTOS, "despesas", orfaos[0]))
        irrf = orfaos[1].valor
        custos.append(_custo("Outras Bovespa", CategoriaCusto.OUTROS, "despesas", orfaos[2]))
    m_dt = RE_IRRF_DT.search(texto)
    if m_dt is None and re.search(r"daytrade.*proje", compactar(texto)):
        raise ExtracaoError("linha de IRRF day trade (projeção) fora do padrão")
    liquido = RE_LIQUIDO.search(texto)
    if not liquido:
        raise ExtracaoError("'Líquido para' não encontrado")
    vlo = _obrigatorio(texto, "Valor líquido das operações")
    return ResumoBovespa(
        valor_operacoes=negs["Valor das operações"],
        compras=negs["Compras à vista"] + negs["Opções - compras"],
        vendas=negs["Vendas à vista"] + negs["Opções - vendas"],
        termo=negs["Operações à termo"],
        valor_liquido_operacoes=vlo,
        custos=tuple(custos),
        total_cblc=_obrigatorio(texto, "Total CBLC"),
        total_bovespa=_obrigatorio(texto, "Total Bovespa / Soma"),
        total_despesas=_obrigatorio(
            texto, "Total Custos / Despesas", "Total corretagem / Despesas"
        ),
        irrf=irrf,
        irrf_base=base,
        irrf_no_liquido=False,  # decidido na conciliação
        irrf_day_trade=decimal_br(m_dt["valor"]) if m_dt else Decimal(0),
        liquido=Montante(valor=decimal_br(liquido["valor"]), dc=liquido["dc"]),  # type: ignore[arg-type]
        data_liquidacao=data_br(liquido["data"]),
    )
