"""Notas extraídas (F2) -> candidato no formato do gabarito, para `validar-gabarito`.

Custos da nota são distribuídos pelas operações na proporção do valor, e o IRRF só pelas vendas
(compra não tem IRRF), para que a comparação no nível "notas" (somas por nota) funcione; não é o
rateio fiscal (F4). O 0,005% vai para as vendas comuns sem exercício nem PME (como a base do motor;
a PME aqui só pela data da venda, sem o lote) e o 1% para as vendas de day trade. No nível
"negocios" só quantidade e valor de cada (nota, papel, C/V) são comparados.
"""

from collections.abc import Mapping, Sequence
from decimal import Decimal
from pathlib import Path

from pydantic import BaseModel

from coir.domain.nota import NotaExtraida, OperacaoExtraida
from coir.domain.parametros import IsencaoPME
from coir.gabarito.exportar import (
    ARQUIVOS,
    exigir_fora_do_repositorio,
    gravar_csv,
    verificar_pseudonimizacao,
)
from coir.gabarito.modelo import Categoria, Gabarito, OperacaoGabarito
from coir.gabarito.pseudonimos import MapaPseudonimos

OITO_CASAS = Decimal("0.00000001")
EXERCICIO = frozenset({"EXERC CALL", "EXERC PUT"})  # como MERCADOS_EXERCICIO do motor


def _distribuir(valor: Decimal, pesos: Sequence[Decimal]) -> list[Decimal]:
    """Parcelas proporcionais aos pesos que somam exatamente o valor (sobra na última)."""
    total = sum(pesos, Decimal(0))
    if not pesos:
        return []
    if total == 0:
        return [valor if i == len(pesos) - 1 else Decimal(0) for i in range(len(pesos))]
    partes = [(valor * p / total).quantize(OITO_CASAS) for p in pesos[:-1]]
    return [*partes, valor - sum(partes, Decimal(0))]


def _irrf(
    nota: NotaExtraida, ops: list[OperacaoExtraida], pesos: list[Decimal],
    pme: Mapping[str, IsencaoPME],
) -> list[Decimal]:  # fmt: skip
    """IRRF da nota por operação: o 0,005% nas vendas comuns tributáveis, o 1% nas de day trade."""
    resumo = nota.resumo_bovespa or nota.resumo_bmf
    assert resumo is not None

    def isenta(o: OperacaoExtraida) -> bool:
        p = pme.get(o.papel)
        return o.mercado in EXERCICIO or (p is not None and o.data <= p.vendas_ate)

    zero = Decimal(0)
    vendas = [p if o.cv == "V" else zero for o, p in zip(ops, pesos, strict=True)]
    dt = [p if o.categoria == "DAY_TRADE" else zero for o, p in zip(ops, vendas, strict=True)]
    comuns = [zero if d or isenta(o) else p for o, p, d in zip(ops, vendas, dt, strict=True)]
    partes: list[list[Decimal]] = []
    for valor, base in ((resumo.irrf, comuns), (resumo.irrf_day_trade, dt)):
        usada = next((b for b in (base, vendas, pesos) if any(b)), pesos)
        partes.append(_distribuir(valor, usada) if valor else [zero] * len(ops))
    return [a + b for a, b in zip(*partes, strict=True)]


def montar(
    notas: Sequence[tuple[NotaExtraida, list[OperacaoExtraida]]],
    mapa: MapaPseudonimos,
    titular: str = "TITULAR-1",
    pme: Mapping[str, IsencaoPME] | None = None,
) -> Gabarito:
    linhas: list[OperacaoGabarito] = []
    for nota, ops in notas:
        pesos = [abs(o.total) for o in ops]
        custos = _distribuir(nota.custos_total, pesos)
        irrf = _irrf(nota, ops, pesos, pme or {})
        for o, c, i in zip(ops, custos, irrf, strict=True):
            linhas.append(
                OperacaoGabarito(
                    categoria=Categoria(o.categoria),
                    linha=len(linhas) + 2,
                    titular=titular,
                    corretora=o.corretora,
                    conta=mapa.conta(o.corretora, o.conta),
                    nota=o.nota,
                    data=o.data,
                    cv=o.cv,
                    papel=o.papel,
                    mercado=o.mercado,
                    preco=o.preco,
                    quantidade=o.quantidade,
                    total=o.total,
                    custos=c,
                    pm=None,
                    irrf=i,
                    prazo=None,
                    ativo_objeto=o.ativo_objeto,
                )
            )
    return Gabarito(operacoes=linhas, eventos=[], eventos_cadastro=[], posicoes=[], apuracao=[],
                    darfs=[], isentos=[], custodia_inicial=[], avisos=[])  # fmt: skip


def gravar(candidato: Gabarito, mapa: MapaPseudonimos, destino: Path) -> None:
    """Grava os 8 CSV no formato do gabarito (pseudonimizados), fora do repositório."""
    exigir_fora_do_repositorio(destino)
    destino.mkdir(parents=True, exist_ok=True)
    for campo, (nome, modelo) in ARQUIVOS.items():
        registros: list[BaseModel] = getattr(candidato, campo)
        gravar_csv(destino / nome, registros, modelo)
    verificar_pseudonimizacao(destino, mapa)
