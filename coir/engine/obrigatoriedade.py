"""Avisos da declaração anual: obrigatoriedade pelos dados de bolsa e total informativo do IR
mínimo (Lei 9.250 art. 16-A).

Obrigatoriedade (decisão do usuário de
10/11/2026, leitura conservadora). Os critérios vêm de data/parametros_fiscais.csv, pelo
ano-calendário (ficha docs/normas/in-rfb-dirpf-anuais.md):

- até o ano-calendário 2021 (exercício 2022): qualquer operação em bolsa obriga;
- desde o ano-calendário 2022 (exercício 2023): alienações em bolsa acima de R$ 40 mil no ano ou
  ganho líquido sujeito ao imposto em algum mês.

O COIR só olha a bolsa; os demais critérios (rendimentos, bens) ficam com o usuário.
"""

from collections.abc import Iterable
from datetime import date
from decimal import Decimal

from coir.domain.fiscal import ZERO, MesApurado, Movimento
from coir.domain.formato import reais
from coir.domain.parametros import ParametrosFiscais


def _vigente(params: ParametrosFiscais, nome: str, dia: date) -> str | None:
    achados = [p for p in params.parametros if p.nome == nome and p.vigente(dia)]
    return achados[0].valor if len(achados) == 1 else None


def aviso_obrigatoriedade(
    titular: int,
    ano: int,
    movimentos: Iterable[Movimento],
    meses: Iterable[MesApurado],
    params: ParametrosFiscais,
) -> str | None:
    fim = date(ano, 12, 31)
    criterio = _vigente(params, "obrigatoriedade_bolsa_criterio", fim)
    do_ano = [m for m in movimentos if m.titular == titular and m.data.year == ano
              and m.origem.startswith("op:")]  # fmt: skip
    exercicio = ano + 1
    if criterio == "qualquer_operacao":
        if do_ano:
            return (f"Declaração do exercício {exercicio} obrigatória: houve operação em bolsa em "
                    f"{ano} (IN RFB da DIRPF {exercicio}, art. 2º III)")  # fmt: skip
        return None
    if criterio != "vendas_ou_ganho":
        return None
    limite = Decimal(_vigente(params, "obrigatoriedade_bolsa_vendas", fim) or "0")
    vendas = sum((m.valor for m in do_ano if m.cv == "V"), ZERO)
    ganho = any(r.imposto > 0 or r.resultado > 0 for m in meses
                if m.titular == titular and m.ano == ano for r in m.categorias)  # fmt: skip
    motivos: list[str] = []
    if vendas > limite:
        motivos.append(f"vendas em bolsa de {reais(vendas)} no ano, acima de {reais(limite)}")
    if ganho:
        motivos.append("ganho líquido sujeito ao imposto em algum mês")
    if not motivos:
        return None
    return (f"Declaração do exercício {exercicio} obrigatória pelos dados de bolsa: "
            + " e ".join(motivos) + f" (IN RFB da DIRPF {exercicio}, art. 2º IV)")  # fmt: skip


def aviso_irpf_minimo(
    titular: int, ano: int, meses: Iterable[MesApurado], params: ParametrosFiscais
) -> str | None:
    """Total informativo dos ganhos líquidos em bolsa que entram na base do IR mínimo (Lei
    9.250/1995 art. 16-A §1º I, incluído pela Lei 15.270/2025, desde o ano-calendário 2026): os
    ganhos tributados (resultado positivo do mês, por categoria) e os isentos (ações até o limite e
    PME); o imposto apurado sobre eles é dedutível (§3º IV). Decisão do usuário de 10/11/2026: só
    informa, o COIR não calcula o IR mínimo (fora do escopo, R14)."""
    renda = _vigente(params, "irpf_minimo_renda_anual", date(ano, 12, 31))
    if renda is None:
        return None
    do_ano = [m for m in meses if m.titular == titular and m.ano == ano]
    tributado = sum((r.resultado for m in do_ano for r in m.categorias if r.resultado > 0), ZERO)
    isento = sum((max(r.isento, ZERO) for m in do_ano for r in m.categorias), ZERO)
    isento += sum((m.isento_pme for m in do_ano), ZERO)
    if not tributado and not isento:
        return None
    imposto = sum((m.imposto for m in do_ano), ZERO)
    return (f"IR mínimo (Lei 9.250/1995 art. 16-A, desde 2026), só para quem recebe mais de "
            f"{reais(Decimal(renda))} no ano: os ganhos líquidos em bolsa de {ano} entram na base, "
            f"{reais(tributado + isento)} (tributados {reais(tributado)} e isentos "
            f"{reais(isento)}); o imposto apurado sobre eles, {reais(imposto)}, é dedutível (§3º "
            "IV). Total informativo: o COIR não calcula o IR mínimo")  # fmt: skip
