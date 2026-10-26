"""Futuros (índice, dólar, commodities): resultado no encerramento da posição.

Lei 11.051/2004 art. 32 (caput e §2º II): os resultados em mercados de liquidação futura sujeitos a
ajustes "serão reconhecidos por ocasião da liquidação do contrato, cessão ou encerramento da
posição". IN RFB 1.585/2015 art. 61 e P&R IRPF 2026 pergunta 728: resultado de cada contrato = soma
algébrica dos ajustes diários entre a abertura e o encerramento ou liquidação, total ou parcial.

Implementação: por titular e contrato (ex.: WINM24), os ajustes da posição aberta (AJUPOS e o ajuste
do dia do negócio que abriu) e os custos rateados da abertura ficam acumulados; no encerramento
parcial sai a parte proporcional (preço médio, decisão do usuário de 15/10/2026), somada ao ajuste
e aos custos do negócio que encerra. No vencimento, sai o que sobrar. Day trade (IN 1.585 art. 65
§§10 e 11) e taxas sem negócio (permanência) têm resultado no próprio dia.
"""

from collections import defaultdict
from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

from coir.domain.fiscal import (
    ZERO,
    AjusteFuturo,
    Categoria,
    Grupo,
    Movimento,
    NotaMotor,
    PosicaoFuturo,
    TipoAjuste,
    TipoAtivo,
)
from coir.engine.rateio import ratear
from coir.erros import MotorError


@dataclass
class _Contrato:
    grupo: Grupo
    quantidade: Decimal = ZERO  # contratos: comprado > 0, vendido < 0
    acumulado: Decimal = ZERO  # ajustes e custos da posição aberta, ainda não realizados
    vencimento: date | None = None
    entrega_fisica: bool = False
    origens: set[tuple[int, str]] = field(default_factory=lambda: set[tuple[int, str]]())

    def origem(self) -> tuple[int | None, str | None]:
        """Conta e corretora da posição (None quando está em mais de uma)."""
        if len(self.origens) == 1:
            return next(iter(self.origens))
        corretoras = {c for _, c in self.origens}
        return (None, next(iter(corretoras)) if len(corretoras) == 1 else None)


def movimentos_futuros(
    notas: Iterable[NotaMotor], ate: date, avisos: list[str] | None = None,
    abertas: list[PosicaoFuturo] | None = None,
) -> list[Movimento]:  # fmt: skip
    """Resultados realizados até `ate`; `abertas` recebe os contratos ainda em aberto."""
    saida: list[Movimento] = []
    contratos: dict[tuple[int, str], _Contrato] = {}
    por_dia: dict[date, list[NotaMotor]] = defaultdict(list)
    for nota in notas:
        if nota.data > ate:
            continue
        if nota.ajustes:
            por_dia[nota.data].append(nota)
        elif not nota.operacoes and nota.custos:  # custos sem linha para ratear: não somem
            raise MotorError(f"nota {nota.numero} de {nota.data}: custos de {nota.custos} sem "
                             "negócio nem ajuste para ratear")  # fmt: skip
    vencimentos = {a.vencimento for n in (x for d in por_dia.values() for x in d)
                   for a in n.ajustes if a.vencimento and a.vencimento <= ate}  # fmt: skip
    for dia in sorted(set(por_dia) | vencimentos):
        linhas: list[tuple[NotaMotor, AjusteFuturo, Decimal]] = []
        for nota in sorted(por_dia.get(dia, []), key=lambda n: (n.numero.zfill(20), n.id)):
            custos = ratear(nota.custos, [a.volume for a in nota.ajustes])
            linhas += [(nota, a, c) for a, c in zip(nota.ajustes, custos, strict=True)]
        # ajuste da posição carregada antes dos negócios do dia, em todas as notas (contas) do dia
        for nota, a, custo in sorted(linhas, key=lambda x: x[1].tipo is not TipoAjuste.AJUPOS):
            saida += _linha(contratos, nota, a, custo, avisos)
        for (titular, mercadoria), c in list(contratos.items()):
            if c.quantidade and c.vencimento == dia:
                if c.entrega_fisica and avisos is not None:
                    avisos.append(f"{dia}: {mercadoria} com {c.quantidade} contrato(s) aberto(s) "
                                  "no vencimento; a liquidação por entrega física está fora do "
                                  "escopo: o resultado é o dos ajustes acumulados")  # fmt: skip
                saida.append(_mov(dia, titular, c.origem(), mercadoria, c.grupo, Categoria.COMUM,
                                  "futuro: vencimento do contrato, resultado = ajustes acumulados",
                                  c.acumulado, ZERO, "vencimento"))  # fmt: skip
                c.quantidade, c.acumulado = ZERO, ZERO
                c.origens.clear()
    for (titular, mercadoria), c in sorted(contratos.items()):
        if c.quantidade and abertas is not None:
            abertas.append(PosicaoFuturo(titular=titular, mercadoria=mercadoria, grupo=c.grupo,
                                         quantidade=c.quantidade, acumulado=c.acumulado,
                                         vencimento=c.vencimento))  # fmt: skip
        if c.quantidade and c.vencimento is None and avisos is not None:
            avisos.append(f"{mercadoria}: posição de {c.quantidade} contrato(s) sem vencimento "
                          "conhecido; o resultado só sai no encerramento")  # fmt: skip
    return saida


def _linha(
    contratos: dict[tuple[int, str], _Contrato], nota: NotaMotor, a: AjusteFuturo, custo: Decimal,
    avisos: list[str] | None,
) -> list[Movimento]:  # fmt: skip
    origem = f"nota:{nota.id}"
    de = (nota.conta, nota.corretora)
    if a.tipo in (TipoAjuste.DAY_TRADE, TipoAjuste.OUTRO):
        categoria = Categoria.DAY_TRADE if a.day_trade else Categoria.COMUM
        regra = "futuro: day trade, resultado = ajuste - custos" if a.day_trade else (
            "futuro: taxa sem negócio, resultado no dia")  # fmt: skip
        return [_mov(nota.data, nota.titular, de, a.mercadoria, a.grupo, categoria, regra,
                     a.valor - custo, custo, origem)]  # fmt: skip
    chave = (nota.titular, a.mercadoria)
    c = contratos.setdefault(chave, _Contrato(a.grupo))
    c.vencimento = a.vencimento or c.vencimento
    c.entrega_fisica = c.entrega_fisica or a.entrega_fisica
    if a.tipo is TipoAjuste.AJUPOS:
        if c.quantidade == 0:  # posição de antes do livro: não há abertura para diferir
            if avisos is not None:
                avisos.append(f"{nota.data}: ajuste de posição de {a.mercadoria} sem posição "
                              "aberta no livro; resultado no dia")  # fmt: skip
            return [_mov(nota.data, nota.titular, de, a.mercadoria, a.grupo,
                         Categoria.COMUM, "futuro: ajuste sem posição no livro (resultado no dia)",
                         a.valor - custo, custo, origem)]  # fmt: skip
        c.acumulado += a.valor - custo
        return []
    sinal = Decimal(1) if a.cv == "C" else Decimal(-1)
    q = a.quantidade
    if c.quantidade == 0 or (c.quantidade > 0) == (sinal > 0):
        c.origens.add(de)
        c.quantidade += sinal * q
        c.acumulado += a.valor - custo
        return []
    fecha = min(q, abs(c.quantidade))
    parte = c.acumulado if fecha == abs(c.quantidade) else c.acumulado * fecha / abs(c.quantidade)
    proprio = (a.valor - custo) * fecha / q
    resultado = parte + proprio
    c.acumulado -= parte
    c.quantidade += sinal * fecha
    if c.quantidade == 0:
        c.acumulado = ZERO
        c.origens.clear()
    sobra = q - fecha
    if sobra:  # inverte a posição: o resto do negócio abre no outro sentido
        c.origens.add(de)
        c.quantidade += sinal * sobra
        c.acumulado += (a.valor - custo) - proprio
    return [_mov(nota.data, nota.titular, de, a.mercadoria, a.grupo, Categoria.COMUM,
                 "futuro: encerramento (Lei 11.051/2004 art. 32), resultado = ajustes acumulados "
                 "pelo preço médio + ajuste do negócio - custos", resultado,
                 custo * fecha / q, origem, fecha)]  # fmt: skip


def _mov(
    dia: date, titular: int, de: tuple[int | None, str | None], mercadoria: str, grupo: Grupo,
    categoria: Categoria, regra: str, resultado: Decimal, custos: Decimal, origem: str,
    quantidade: Decimal = ZERO,
) -> Movimento:  # fmt: skip
    return Movimento(
        data=dia, titular=titular, conta=de[0], corretora=de[1], ativo=None, ticker=mercadoria,
        origem=origem,
        regra=regra, categoria=categoria, grupo=grupo, tipo=TipoAtivo.FUTURO, cv=None,
        quantidade=quantidade, valor=resultado + custos, custos=custos, resultado=resultado,
        quantidade_apos=ZERO, custo_apos=ZERO,
    )  # fmt: skip
