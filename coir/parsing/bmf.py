"""Gramática BM&F (SINACOR) nas variantes XP (antiga e nova) e BTG.

XP: resumo em quatro pares "linha de rótulos" / "linha de valores", com "|" e D/C em posições
variáveis. BTG: cada valor vem antes do seu rótulo; os indicadores D/C ficam soltos no fim e o sinal
é deduzido dos negócios.
"""

import re
from decimal import Decimal
from typing import Literal

from coir.domain.nota import CategoriaCusto, Custo, Montante, NegocioBMF, ResumoBMF
from coir.erros import ExtracaoError
from coir.parsing.bovespa import RE_SO_VALOR
from coir.parsing.numeros import RE_DECIMAL, RE_INTEIRO, compactar, data_br, decimal_br, montantes

Variante = Literal["xp", "btg"]

RE_NEGOCIO = re.compile(
    rf"^\s*(?P<cv>[CV])\s+(?P<mercadoria>[A-Z0-9]{{3,}}(?:\s(?:[FGHJKMNQUVXZ]\d{{2}}|PPPP))?)\s+"
    rf"(?:(?P<contraparte>@)?(?P<vencimento>\d{{2}}/\d{{2}}/\d{{4}})\s+)?"
    rf"(?P<qtd>{RE_INTEIRO})\s+(?P<preco>{RE_DECIMAL})\s+(?P<tipo>[A-ZÀ-Ú][A-ZÀ-Ú. ]*?)\s+"
    rf"(?P<valor>{RE_DECIMAL})(?:\s+(?P<dc>[DC]))?\s+(?P<taxa>{RE_DECIMAL})\s*$"
)
# XP: rótulo compacto que abre cada linha de rótulos -> nomes dos valores da linha seguinte
LINHAS_XP: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("vendadisponivel", ("venda_disp", "compra_disp", "venda_opc", "compra_opc", "valor_negocios")),
    ("irrfirrfday", ("irrf", "irrf_dt", "taxa_operacional", "taxa_registro", "taxas_bmf")),
    ("+outroscustos", ("outros_custos", "impostos", "ajuste_posicao", "ajuste_dt", "total_custos")),
    ("outrosirrfoperacional", ("outros", "irrf_operacional", "conta_investimento", "conta_normal",
                               "liquido_parcial", "liquido")),
)  # fmt: skip
# BTG: rótulos na ordem do texto; o i-ésimo valor do bloco pertence ao i-ésimo rótulo
ROTULOS_BTG: tuple[tuple[str, str], ...] = (
    ("outros", "Outros"),
    ("irrf_corretagem", "IRRF Corretagem"),
    ("conta_investimento", "Total Conta Investimento"),
    ("conta_normal", "Total Conta Normal"),
    ("liquido_parcial", "Total liquido (#)"),
    ("liquido", "Total líquido da nota"),
    ("ajuste_posicao", "Ajuste de posição"),
    ("ajuste_dt", "Ajuste day trade"),
    ("total_custos", "Total das despesas"),
    ("irrf_dt", "IRRF Day Trade (proj.)"),
    ("taxa_operacional", "Taxa operacional"),
    ("taxa_registro", "Taxa registro BM&F"),
    ("taxas_bmf", "Taxas BM&F (emol+f.gar)"),
    ("irrf", "IRRF"),
    ("compra_disp", "Compra disponível"),
    ("venda_opc", "Venda Opções"),
    ("compra_opc", "Compra Opções"),
    ("valor_negocios", "Valor dos negócios"),
    ("venda_disp", "Venda disponível"),
    ("outros_custos", "+Outros Custos"),
    ("impostos", "I.S.S"),
)


RE_PARECE_NEGOCIO = re.compile(r"^\s*[CV]\s+[A-Z0-9]{3,}\b.*\d,\d")


def negocios(texto: str) -> list[NegocioBMF]:
    saida: list[NegocioBMF] = []
    for linha in texto.splitlines():
        m = RE_NEGOCIO.match(linha)
        if not m:
            if RE_PARECE_NEGOCIO.match(linha):
                raise ExtracaoError(f"linha de negócio BM&F fora do padrão: {linha.strip()!r}")
            continue
        venc = m["vencimento"]
        saida.append(
            NegocioBMF(
                cv=m["cv"],  # type: ignore[arg-type]
                mercadoria=m["mercadoria"].replace(" ", ""),
                vencimento=data_br(venc) if venc and not venc.endswith("/0001") else None,
                quantidade=decimal_br(m["qtd"]),
                preco=decimal_br(m["preco"]),
                tipo=re.sub(r"\s+", " ", m["tipo"]).strip(),
                valor=decimal_br(m["valor"]),
                dc=m["dc"],  # type: ignore[arg-type]
                taxa_operacional=decimal_br(m["taxa"]),
            )
        )
    return saida


def _valores_xp(texto: str) -> dict[str, Montante]:
    linhas = [linha for linha in texto.splitlines() if linha.strip()]
    valores: dict[str, Montante] = {}
    for inicio, nomes in LINHAS_XP:
        i = next((k for k, linha in enumerate(linhas) if compactar(linha).startswith(inicio)), None)
        if i is None or i + 1 >= len(linhas):
            raise ExtracaoError(f"resumo BM&F: linha de rótulos {inicio!r} não encontrada")
        lidos = montantes(linhas[i + 1])
        if len(lidos) != len(nomes):
            raise ExtracaoError(
                f"resumo BM&F: {len(lidos)} valores após {inicio!r} (esperado {len(nomes)})"
            )
        valores.update(zip(nomes, lidos, strict=True))
    return valores


def _valores_btg(texto: str, ajuste_dt: Decimal, total: Decimal) -> dict[str, Montante]:
    linhas = [linha.strip() for linha in texto.splitlines() if linha.strip()]
    compactos = [compactar(linha) for linha in linhas]
    try:
        inicio = compactos.index("outros")
        fim = compactos.index("i.s.s", inicio)
    except ValueError as erro:
        raise ExtracaoError("resumo BM&F (BTG): bloco Outros ... I.S.S não encontrado") from erro
    valores = [v for linha in linhas[inicio:fim] if RE_SO_VALOR.match(linha) or _so_valores(linha)
               for v in montantes(linha)]  # fmt: skip
    if len(valores) != len(ROTULOS_BTG):
        raise ExtracaoError(
            f"resumo BM&F (BTG): {len(valores)} valores (esperado {len(ROTULOS_BTG)})"
        )
    lidos = dict(zip((nome for nome, _ in ROTULOS_BTG), valores, strict=True))
    # D/C soltos no fim: o sinal de cada ajuste vem dos negócios (a conciliação confere os valores)
    sinais = {"valor_negocios": total, "ajuste_dt": ajuste_dt, "ajuste_posicao": total - ajuste_dt}
    for nome, referencia in sinais.items():
        lidos[nome] = Montante(valor=lidos[nome].valor, dc="C" if referencia >= 0 else "D")
    return lidos


def _so_valores(linha: str) -> bool:
    return bool(re.fullmatch(rf"(?:{RE_DECIMAL}\s*)+", linha))


def resumo(texto: str, variante: Variante, negocios_bmf: list[NegocioBMF]) -> ResumoBMF:
    if variante == "xp":
        v = _valores_xp(texto)
    else:

        def soma(so_day_trade: bool) -> Decimal:
            return sum((n.valor if n.dc == "C" else -n.valor for n in negocios_bmf
                        if n.day_trade or not so_day_trade), Decimal(0))  # fmt: skip

        v = _valores_btg(texto, soma(True), soma(False))
    custos = (
        Custo(rotulo="Taxa operacional", categoria=CategoriaCusto.CORRETAGEM, grupo="bmf",
              valor=v["taxa_operacional"].valor),
        Custo(rotulo="Taxa registro BM&F", categoria=CategoriaCusto.REGISTRO, grupo="bmf",
              valor=v["taxa_registro"].valor),
        Custo(rotulo="Taxas BM&F (emol+f.gar)", categoria=CategoriaCusto.TAXAS_BMF, grupo="bmf",
              valor=v["taxas_bmf"].valor),
        Custo(rotulo="+Outros Custos", categoria=CategoriaCusto.OUTROS, grupo="bmf",
              valor=v["outros_custos"].valor),
        Custo(rotulo="Outros", categoria=CategoriaCusto.OUTROS, grupo="bmf",
              valor=v["outros"].valor),
    )  # fmt: skip
    liquido = v["liquido"]
    if variante == "btg":
        base = v["ajuste_dt"].assinado - v["total_custos"].valor - v["impostos"].valor
        liquido = Montante(valor=liquido.valor, dc="C" if base >= 0 else "D")
    return ResumoBMF(
        valor_negocios=v["valor_negocios"],
        ajuste_posicao=v["ajuste_posicao"],
        ajuste_day_trade=v["ajuste_dt"],
        custos=custos,
        total_custos=v["total_custos"].valor,
        impostos=v["impostos"].valor,
        irrf=v["irrf"].valor,
        irrf_day_trade=v["irrf_dt"].valor,
        liquido=liquido,
    )
