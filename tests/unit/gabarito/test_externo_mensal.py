"""Relatório mensal de outro sistema (texto sintético) e comparação com a apuração do motor."""

from decimal import Decimal
from pathlib import Path

import pytest

from coir.domain.fiscal import Categoria, Ficha, FichaMes, Grupo, MesApurado, ResultadoCategoria
from coir.erros import GabaritoError
from coir.gabarito.externo_mensal import (
    Explicacao,
    carregar_explicacoes,
    comparar_mensal,
    explicar,
    ler_texto,
    saldo_inicial,
)

D = Decimal
MESES = ("Janeiro", "Fevereiro")


def _pagina(mes: int, acoes: str = "R$ 1.000,00", anterior: str = "R$ 5.000,00 R$ 300,00",
            prejuizo: str = "R$ 4.000,00 R$ 300,00", extra: str = "") -> str:  # fmt: skip
    linhas = [
        "Operações Comuns e Day-Trade TITULAR DE TESTE",
        f"{MESES[mes - 1]} 2020 Mesma ordem e nomenclatura do programa",
        "Mercado à Vista Operações Comuns Day-Trade",
        f"Mercado à vista - ações {acoes} R$ 0,00",
        *(f"{r} R$ 0,00 R$ 0,00" for r in (
            "Mercado à vista - ouro", "Mercado à vista - ouro at. fin. fora bolsa",
            "Mercado opções - ações", "Mercado opções - ouro", "Mercado opções - fora de bolsa",
            "Mercado opções - outros", "Mercado futuro - dólar dos EUA",
            "Mercado futuro - índices", "Mercado futuro - juros")),
        "Mercado futuro - outros R$ 0,00 -R$ 50,00",
        "Mercado a termo - ações/ouro R$ 0,00 R$ 0,00",
        "Mercado a termo - outros R$ 0,00 R$ 0,00",
        f"Resultado líquido do mês {acoes} R$ 0,00",
        f"Resultado negativo até o mês anterior {anterior}",
        "Base de cálculo do imposto R$ 0,00 R$ 0,00",
        f"Prejuízo a compensar {prejuizo}",
        "Alíquota do imposto 15% 20%",
        "Imposto devido R$ 0,00 R$ 0,00",
        "Total do imposto devido R$ 0,00",
        *(f"{r} R$ 0,00" for r in (
            "IR fonte de Day Trade no mês", "IR fonte de Day Trade nos meses anteriores",
            "IR fonte de Day Trade a compensar")),
        "IR fonte (Lei nr 11.033/2004) R$ 1,10",
        "IR fonte (Lei nr 11.033/2004) nos meses anteriores R$ 0,00",
        "IR fonte (Lei nr 11.033/2004) a compensar R$ 1,10",
        "Imposto a pagar R$ 0,00",
        "Total em alienações de ações (op. comuns) (Lei nr 11.033/2004) R$ 22.000,00",
        *(f"{r} R$ 0,00" for r in (
            "Lucro isento (Lei nr 11.033/2004)", "Lucro isento (Lei nr 11.478/2007)",
            "Lucro isento (Lei nr 13.043/2014)", "Lucro isento em ouro - ativo financeiro")),
        extra,
    ]  # fmt: skip
    return "\n".join(x for x in linhas if x)


def _apurado(mes: int, acoes: str) -> MesApurado:
    a_vista, outros = D(acoes), D(0)
    comum = ResultadoCategoria(
        categoria=Categoria.COMUM, por_grupo={Grupo.A_VISTA: a_vista, Grupo.COMMODITY: outros},
        resultado=a_vista + outros, vendas_acoes=D("22000"), prejuizo_anterior=D("-5000"),
        prejuizo_a_compensar=D("-5000") + a_vista + outros, aliquota=D("0.15"),
    )  # fmt: skip
    dt = ResultadoCategoria(categoria=Categoria.DAY_TRADE, por_grupo={Grupo.COMMODITY: D(-50)},
                            resultado=D(0),
                            prejuizo_anterior=D("-300"), prejuizo_a_compensar=D("-300"),
                            aliquota=D("0.2"))  # fmt: skip
    fii = ResultadoCategoria(categoria=Categoria.FII, por_grupo={}, resultado=D(0),
                             aliquota=D("0.2"))  # fmt: skip
    ficha = FichaMes(ficha=Ficha.COMUNS, imposto=D(0), irrf_mes=D("1.10"),
                     irrf_a_compensar=D("1.10"))  # fmt: skip
    return MesApurado(
        titular=1, ano=2020, mes=mes, categorias=(comum, dt, fii), imposto=D(0),
        irrf_mes=D("1.10"), irrf_saldo_anterior=D(0), irrf_deduzido=D(0), irrf_saldo=D("1.10"),
        imposto_devido=D(0), darf_acumulado_anterior=D(0), darf=D(0), darf_acumulado=D(0),
        vencimento=None, fichas=(ficha,),
    )  # fmt: skip


def test_le_os_meses_com_negativos_e_duas_colunas() -> None:
    capa = "APURAÇÃO DE IMPOSTO\nCLIENTE\nTITULAR DE TESTE"
    meses = ler_texto([capa, _pagina(1), _pagina(2, acoes="-R$ 1.234,56")])
    assert [(m.ano, m.mes) for m in meses] == [(2020, 1), (2020, 2)]
    assert meses[1].valor("Mercado à vista - ações") == D("-1234.56")
    assert meses[0].valor("Mercado futuro - outros", 1) == D("-50")
    assert meses[0].valores["Imposto a pagar"] == (D(0), None)
    assert meses[0].valor("Alíquota do imposto", 1) == D("0.2")


def test_saldo_inicial_e_o_prejuizo_recebido_em_janeiro() -> None:
    inicial = saldo_inicial(ler_texto([_pagina(1, anterior="R$ 111.174,75 R$ 11.684,22")]), 7)
    assert (inicial.titular, inicial.ano) == (7, 2019)
    assert (inicial.comum, inicial.day_trade) == (D("111174.75"), D("11684.22"))
    with pytest.raises(GabaritoError, match="janeiro"):
        saldo_inicial(ler_texto([_pagina(2)]), 1)


def test_linha_desconhecida_ou_faltando_falha_alto() -> None:
    with pytest.raises(GabaritoError, match="desconhecida"):
        ler_texto([_pagina(1, extra="Mercado novo - cripto R$ 1,00 R$ 0,00")])
    sem_linha = "\n".join(x for x in _pagina(1).splitlines() if not x.startswith("Imposto a"))
    with pytest.raises(GabaritoError, match="faltam"):
        ler_texto([sem_linha])
    with pytest.raises(GabaritoError, match="nenhum mês"):
        ler_texto(["capa sem mês"])


def test_compara_linha_a_linha_e_explica_com_curinga(tmp_path: Path) -> None:
    externos = ler_texto([_pagina(1), _pagina(2, acoes="R$ 900,00",
                                              prejuizo="R$ 4.100,00 R$ 300,00")])  # fmt: skip
    linhas = comparar_mensal(externos, [_apurado(1, "1000"), _apurado(2, "1000")], 1, D("1"))
    abertas = {(x.mes, x.rotulo, x.coluna): x.diferenca for x in linhas if x.classe != "confere"}
    assert abertas == {
        (2, "Mercado à vista - ações", "comuns"): D(100),
        (2, "Resultado líquido do mês", "comuns"): D(100),
        (2, "Prejuízo a compensar", "comuns"): D(-100),
    }
    arquivo = tmp_path / "explicadas.csv"
    arquivo.write_text("ano,mes,rotulo,coluna,motivo\n"
                       "2020,2,Mercado à vista - ações,comuns,PM inicial errado no outro sistema\n"
                       "2020,*,*,comuns,efeito no resultado e no prejuízo\n",
                       encoding="utf-8")  # fmt: skip
    explicadas = explicar(linhas, carregar_explicacoes(arquivo))
    assert {x.classe for x in explicadas} == {"confere", "explicada"}
    assert next(x.motivo for x in explicadas if x.rotulo == "Mercado à vista - ações"
                and x.mes == 2) == "PM inicial errado no outro sistema"  # fmt: skip
    so_mes_1 = [Explicacao(2020, "1", "*", "*", "outro mês")]
    assert all(x.classe != "explicada" for x in explicar(linhas, so_mes_1))
