"""Piloto do gabarito: põe lado a lado, para poucas notas, o que o gabarito registra e o resumo
financeiro impresso no PDF. Serve para o usuário conferir o método antes de gerar tudo.

A leitura do PDF aqui é propositalmente simples (linhas do resumo + "Líquido para"); o extrator de
verdade é a fase F2.
"""

import re
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

from coir.erros import GabaritoError
from coir.gabarito.modelo import Gabarito, OperacaoGabarito
from coir.gabarito.normalizar import ALIAS_CORRETORA
from coir.parsing.texto_pdf import paginas

NOTAS_PADRAO = (
    "Agora_NC_2008_04.pdf",
    "Rico_NC_2015_08.pdf",
    "XP_NC_2018_06.pdf",
    "XP_NC_2026_08.pdf",
    "XP_NC_2019_02_BMF.pdf",
    "XP_NC_2023_06_BMF.pdf",
    "BTG_NC_2023_01.pdf",
    "BTG_NC_2023_08_BMF.pdf",
)
MERCADOS_BMF = {"MILHO", "BOI", "INDICE", "DOLAR", "CAFE", "SOJA"}
RE_ARQUIVO = re.compile(
    r"^(?P<corretora>[A-Za-z]+)_NC_(?P<ano>\d{4})_(?P<mes>\d{2})(?P<bmf>_BMF)?\.pdf$"
)
RE_LIQUIDO = re.compile(
    r"L[ií]quido para \d{2}/\d{2}/\d{4}\s+(?P<valor>[\d.]+,\d{2})\s*\|?\s*(?P<dc>[DC])"
)
RE_RESUMO = re.compile(
    r"(?i)(valor das opera|valor dos neg|vendas? [àa] vista|compras? [àa] vista|op[çc][õo]es|"
    r"taxa|emolumentos|corretagem|custos|total|irrf|i\.r\.r\.f|ajuste|l[ií]quido|impostos|iss)"
)
RE_SENSIVEL = re.compile(
    r"(?i)(cliente|cpf|c\.n\.p\.j|conta|endere|e-mail|telefone|assessor|cep\b)"
)
DUAS = Decimal("0.01")


@dataclass(frozen=True)
class ArquivoNota:
    corretora: str
    ano: int
    mes: int
    bmf: bool


def interpretar_nome(nome: str) -> ArquivoNota:
    m = RE_ARQUIVO.match(nome)
    if not m:
        raise GabaritoError(
            f"Nome de arquivo fora do padrão Corretora_NC_AAAA_MM[_BMF].pdf: {nome}"
        )
    corretora = m["corretora"].upper()
    return ArquivoNota(
        ALIAS_CORRETORA.get(corretora, corretora), int(m["ano"]), int(m["mes"]), bool(m["bmf"])
    )


def _br(valor: Decimal) -> str:
    texto = f"{valor.quantize(DUAS):,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def _numero_br(texto: str) -> Decimal:
    return Decimal(texto.replace(".", "").replace(",", "."))


RE_INTEIRO_IMPRESSO = re.compile(r"(?<![\d.,])\d[\d.]*(?<!\.)(?![\d.,]*\d|,)")


def paginas_da_nota(textos: list[str], nota: str) -> list[int]:
    """Páginas em que o número da nota aparece como inteiro isolado, com ou sem pontos de
    milhar e zeros à esquerda (ex.: 57990, 57.990, 057.990). Valores com vírgula não contam."""
    alvo = int(nota)
    paginas: list[int] = []
    for i, texto in enumerate(textos):
        for token in RE_INTEIRO_IMPRESSO.findall(texto):
            digitos = token.replace(".", "")
            if digitos.isdigit() and int(digitos) == alvo:
                paginas.append(i)
                break
    return paginas


RE_SO_VALORES = re.compile(r"^[\d.,|\sDC+-]*\d,\d{2}[\d.,|\sDC+-]*$")


def linhas_resumo(texto: str) -> list[str]:
    """Linhas do resumo financeiro: rótulos conhecidos e linhas só de valores (layout BM&F, em que
    os valores ficam na linha abaixo dos rótulos)."""
    saida: list[str] = []
    for linha in texto.splitlines():
        linha = linha.strip()
        if not linha or RE_SENSIVEL.search(linha):
            continue
        if RE_RESUMO.search(linha) or RE_SO_VALORES.match(linha):
            saida.append(linha[:160])
    return saida


def liquido_pdf(texto: str) -> Decimal | None:
    """Valor do "Líquido para dd/mm/aaaa" com sinal (+ crédito, - débito), se houver."""
    m = RE_LIQUIDO.search(texto)
    if not m:
        return None
    valor = _numero_br(m["valor"])
    return valor if m["dc"] == "C" else -valor


def conferir_liquido(liquido: Decimal, estimado: Decimal, irrf: Decimal) -> str:
    """Compara o líquido do PDF com o estimado, sem e com o IRRF descontado.

    Descontar ou dispensar o IRRF depende de parâmetro fiscal; aqui só se informa qual das duas
    contas confere.
    """
    if abs(liquido - estimado) <= DUAS:
        return "confere" + (" (IRRF não descontado no líquido)" if irrf else "")
    if irrf and abs(liquido - (estimado - irrf)) <= DUAS:
        return "confere (IRRF descontado no líquido)"
    return f"**diferença {_br(liquido - estimado)}**"


def _ops_do_arquivo(gabarito: Gabarito, arq: ArquivoNota) -> dict[str, list[OperacaoGabarito]]:
    por_nota: dict[str, list[OperacaoGabarito]] = defaultdict(list)
    for op in gabarito.operacoes:
        if (op.corretora, op.data.year, op.data.month) != (arq.corretora, arq.ano, arq.mes):
            continue
        if (op.mercado in MERCADOS_BMF) != arq.bmf:
            continue
        por_nota[op.nota].append(op)
    return dict(sorted(por_nota.items(), key=lambda kv: (kv[1][0].data, kv[0])))


def secao(gabarito: Gabarito, pasta_notas: Path, nome: str) -> list[str]:
    arq = interpretar_nome(nome)
    textos = paginas(pasta_notas / nome)
    notas = _ops_do_arquivo(gabarito, arq)
    tipo = "BM&F" if arq.bmf else "Bovespa"
    md = [f"## {nome} — {arq.corretora} {tipo} {arq.mes:02d}/{arq.ano}", ""]
    md.append(
        f"{len(textos)} página(s) no PDF; {len(notas)} nota(s) no gabarito para este arquivo."
    )
    md.append("")
    if not notas:
        md += ["**Nenhuma operação do gabarito para este arquivo** — conferir.", ""]
    for nota, ops in notas.items():
        compras = sum((o.total for o in ops if o.cv == "C"), Decimal(0))
        vendas = sum((o.total for o in ops if o.cv == "V"), Decimal(0))
        custos = sum((o.custos for o in ops), Decimal(0))
        irrf = sum((o.irrf for o in ops), Decimal(0))
        estimado = vendas - compras - custos
        pags = paginas_da_nota(textos, nota)
        md.append(f"### Nota {nota} — pregão {ops[0].data:%d/%m/%Y} ({ops[0].conta})")
        md.append("")
        md.append("| C/V | Papel | Mercado | Qtde | Preço | Total | Custos | IRRF |")
        md.append("|---|---|---|--:|--:|--:|--:|--:|")
        for o in ops:
            md.append(
                f"| {o.cv} | {o.papel} | {o.mercado} | {o.quantidade.normalize():f} "
                f"| {_br(o.preco)} "
                f"| {_br(o.total)} | {_br(o.custos)} | {_br(o.irrf)} |"
            )
        md.append("")
        md.append(
            f"Gabarito: compras {_br(compras)} · vendas {_br(vendas)} · custos {_br(custos)} · "
            f"IRRF {_br(irrf)} · líquido estimado (vendas - compras - custos) "
            f"**{_br(estimado)}**; descontando o IRRF **{_br(estimado - irrf)}**"
        )
        if not pags:
            md += [
                "",
                "PDF: **número da nota não encontrado no texto** — conferir manualmente.",
                "",
            ]
            continue
        texto = textos[pags[-1]]
        liquido = liquido_pdf(texto)
        if liquido is not None and not arq.bmf:
            marca = conferir_liquido(liquido, estimado, irrf)
            md.append(f"PDF (página {pags[-1] + 1}): líquido {_br(liquido)} → {marca}")
        else:
            md.append(f"PDF (página {pags[-1] + 1}): conferir o resumo abaixo")
        md += ["", "```", *linhas_resumo(texto), "```", ""]
    return md


def gerar(gabarito: Gabarito, pasta_notas: Path, nomes: tuple[str, ...] = NOTAS_PADRAO) -> str:
    md = [
        "# Piloto do gabarito",
        "",
        "Para cada PDF: as operações que o gabarito registra para cada nota e o resumo financeiro "
        "impresso no PDF. Em notas Bovespa, o líquido estimado do gabarito (sem e com o IRRF "
        'descontado) é comparado com o "Líquido para" do PDF. Em BM&F, conferir o resumo '
        "manualmente (ajustes e multiplicadores entram só na F2).",
        "",
    ]
    for nome in nomes:
        md += secao(gabarito, pasta_notas, nome)
    return "\n".join(md)
