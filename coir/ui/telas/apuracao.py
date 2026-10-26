"""Tela 12 — Apuração mensal: o ano numa grade e, ao escolher o mês, a memória de cálculo em
blocos (limite de isenção, cada categoria com a compensação, IRRF e DARF, fichas do IRPF e os
movimentos do mês com a regra aplicada e a nota de origem)."""

from collections.abc import Callable
from datetime import date
from decimal import Decimal

from nicegui import run, ui

from coir.domain.fiscal import (
    Apuracao,
    Categoria,
    Ficha,
    MesApurado,
    Movimento,
    ResultadoCategoria,
)
from coir.erros import CoirError
from coir.reports.fiscais import GRUPOS, MESES, NOME_CATEGORIA, NOME_GRUPO
from coir.reports.modelo import Coluna, TipoColuna, Valor, formatar
from coir.servicos import carteira
from coir.servicos.apuracao import conta_na_isencao
from coir.servicos.estado import Estado
from coir.ui import tabela

T = TipoColuna
ZERO = Decimal(0)
COLUNAS_ANO = (
    Coluna("Mês"),
    Coluna("Comum", T.DINHEIRO),
    Coluna("Day trade", T.DINHEIRO),
    Coluna("FII", T.DINHEIRO),
    Coluna("Vendas de ações", T.DINHEIRO),
    Coluna("Isento", T.DINHEIRO),
    Coluna("Imposto", T.DINHEIRO),
    Coluna("IRRF deduzido", T.DINHEIRO),
    Coluna("Imposto devido", T.DINHEIRO),
    Coluna("DARF", T.DINHEIRO),
    Coluna("Acumulado", T.DINHEIRO),
)
COLUNAS_MOV = (
    Coluna("Data", T.DATA),
    Coluna("Ativo"),
    Coluna("Categoria"),
    Coluna("Regra"),
    Coluna("C/V"),
    Coluna("Qtd.", T.QUANTIDADE),
    Coluna("Valor", T.DINHEIRO),
    Coluna("Custo baixado", T.DINHEIRO),
    Coluna("Resultado", T.DINHEIRO),
    Coluna("Isenção"),
)
CURTA = {Categoria.COMUM: "Comum", Categoria.DAY_TRADE: "Day trade", Categoria.FII: "FII"}
NOME_FICHA = {Ficha.COMUNS: "Operações comuns e day trade", Ficha.FII: "FII e Fiagro"}


def _r(dinheiro: Decimal) -> str:
    return formatar(dinheiro, T.DINHEIRO)


def _cat(m: MesApurado, c: Categoria) -> ResultadoCategoria | None:
    return next((r for r in m.categorias if r.categoria is c), None)


def _resultado(m: MesApurado, c: Categoria) -> Decimal:
    r = _cat(m, c)
    return r.resultado if r else ZERO


def _linha_ano(m: MesApurado) -> tuple[Valor, ...]:
    comum = _cat(m, Categoria.COMUM)
    return (
        MESES[m.mes - 1],
        _resultado(m, Categoria.COMUM),
        _resultado(m, Categoria.DAY_TRADE),
        _resultado(m, Categoria.FII),
        comum.vendas_acoes if comum else ZERO,
        comum.isento if comum else ZERO,
        m.imposto,
        m.irrf_deduzido,
        m.imposto_devido,
        m.darf,
        m.darf_acumulado,
    )


def _total_ano(meses: list[MesApurado]) -> tuple[Valor, ...]:
    def soma(f: Callable[[MesApurado], Decimal]) -> Decimal:
        return sum((f(m) for m in meses), ZERO)

    return (
        "Ano",
        soma(lambda m: _resultado(m, Categoria.COMUM)),
        soma(lambda m: _resultado(m, Categoria.DAY_TRADE)),
        soma(lambda m: _resultado(m, Categoria.FII)),
        soma(lambda m: c.vendas_acoes if (c := _cat(m, Categoria.COMUM)) else ZERO),
        soma(lambda m: c.isento if (c := _cat(m, Categoria.COMUM)) else ZERO),
        soma(lambda m: m.imposto),
        soma(lambda m: m.irrf_deduzido),
        soma(lambda m: m.imposto_devido),
        soma(lambda m: m.darf),
        None,
    )


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Apuração mensal").classes("coir-titulo-pagina")
    ui.label(
        f"O resultado de {estado.ano} por categoria e, ao escolher um mês, a memória de cálculo: "
        "de onde vem cada valor, com a regra aplicada (IN RFB 1.585/2015)."
    ).classes("coir-subtitulo")
    area = ui.column().classes("w-full gap-4")
    memoria = ui.column().classes("w-full gap-3")
    meses: list[MesApurado] = []
    estado_apuracao: list[Apuracao] = []

    async def carregar() -> None:
        with area:
            espera = ui.spinner(size="lg")
        ate = estado.fim_do_ano()
        try:
            apuracao = await run.io_bound(estado.apuracao, ate)
            mapa = await run.io_bound(estado.notas_das_operacoes)
        except CoirError as erro:
            area.clear()
            with area:
                ui.label(f"Não foi possível apurar: {erro}").classes("coir-negativo")
            return
        espera.delete()
        if apuracao is None or mapa is None:
            return
        estado_apuracao[:] = [apuracao]
        meses[:] = [
            m for m in apuracao.meses if m.titular == estado.titular and m.ano == estado.ano
        ]
        with area:
            if not meses:
                ui.label("Nada apurado neste ano.").classes("coir-vazio w-full")
                return
            ui.label(f"Apurado até {ate:%d/%m/%Y}. Clique num mês para ver a memória.").classes(
                "coir-subtitulo text-sm"
            )
            tabela.grade(
                COLUNAS_ANO,
                [_linha_ano(m) for m in meses],
                chaves=list(range(len(meses))),
                total=_total_ano(meses),
                ao_clicar=lambda i: abrir(int(str(i)), mapa),
            ).mark("grade-apuracao")
            for aviso in apuracao.avisos:
                if aviso[:4] == str(estado.ano) or not aviso[:4].isdigit():  # avisos do ano
                    ui.label(f"Atenção: {aviso}").classes("coir-carimbo text-sm")

    def abrir(indice: int, mapa: dict[int, int]) -> None:
        m = meses[indice]
        movs = [
            x
            for x in estado_apuracao[0].movimentos
            if x.titular == m.titular
            and x.data.year == m.ano
            and x.data.month == m.mes
            and (x.resultado or x.cv == "V")
        ]
        memoria.clear()
        with memoria:
            _memoria(estado, m, movs, mapa)

    ui.timer(0.01, carregar, once=True)


def _bloco(titulo: str) -> ui.column:
    with ui.column().classes("coir-folha gap-1 min-w-[300px] flex-1") as bloco:
        ui.label(titulo).classes("coir-titulo text-base")
    return bloco


def _par(rotulo: str, valor: str, classe: str = "") -> None:
    with ui.row().classes("w-full justify-between gap-6 no-wrap"):
        ui.label(rotulo).classes("text-sm")
        ui.label(valor).classes(f"coir-numero text-sm text-right {classe}")


def _memoria(estado: Estado, m: MesApurado, movs: list[Movimento], mapa: dict[int, int]) -> None:
    ui.label(f"Memória de {MESES[m.mes - 1].lower()} de {m.ano}").classes(
        "coir-titulo text-xl mt-2"
    )
    params = estado.fiscais.parametros
    dia = date(m.ano, m.mes, 1)
    limite = params.decimal("limite_isencao_acoes", dia)
    comum = _cat(m, Categoria.COMUM)
    with ui.row().classes("w-full items-stretch gap-3"):
        with _bloco("1. Limite de isenção (ações à vista)"):
            vendas = comum.vendas_acoes if comum else ZERO
            _par("Vendas de ações no mês (todas as corretoras)", _r(vendas))
            _par("Limite", _r(limite))
            if comum and comum.isento:
                _par("Ganho isento", _r(comum.isento), "coir-positivo")
            veredito = (
                "dentro do limite: ganho de ações isento"
                if vendas <= limite
                else ("acima do limite: ganho tributável")
            )
            ui.label(veredito).classes("coir-subtitulo text-xs")
            ui.label("Lei 11.033/2004 art. 3º I; units fora do limite.").classes(
                "coir-subtitulo text-xs"
            )
        for r in m.categorias:
            with _bloco(f"{2 + list(Categoria).index(r.categoria)}. {NOME_CATEGORIA[r.categoria]}"):
                for g in GRUPOS:
                    if r.por_grupo.get(g):
                        _par(NOME_GRUPO[g], _r(r.por_grupo[g]))
                _par(
                    "Resultado do mês", _r(r.resultado), "coir-negativo" if r.resultado < 0 else ""
                )
                _par("Prejuízo de meses anteriores", _r(r.prejuizo_anterior))
                _par("Base de cálculo", _r(r.base))
                _par("Alíquota", formatar(r.aliquota, T.PERCENTUAL))
                _par("Imposto", _r(r.imposto))
                _par("Prejuízo a compensar", _r(r.prejuizo_a_compensar))
        with _bloco("5. IRRF e DARF"):
            _par("Imposto das categorias", _r(m.imposto))
            _par("IRRF retido no mês", _r(m.irrf_mes))
            _par("Saldo de IRRF anterior", _r(m.irrf_saldo_anterior))
            _par("IRRF deduzido", _r(m.irrf_deduzido))
            _par("Saldo de IRRF que passa", _r(m.irrf_saldo))
            _par("Imposto devido", _r(m.imposto_devido))
            _par("DARF acumulada (abaixo do mínimo)", _r(m.darf_acumulado_anterior))
            _par("DARF do mês", _r(m.darf), "coir-carimbo" if m.darf else "")
            _par("Acumula para o mês seguinte", _r(m.darf_acumulado))
            if m.vencimento:
                _par("Vencimento", f"{m.vencimento:%d/%m/%Y}")
    if m.fichas:
        with ui.row().classes("w-full items-stretch gap-3"):
            for f in m.fichas:
                with _bloco(f"Ficha do IRPF: {NOME_FICHA[f.ficha]}"):
                    _par("Imposto devido", _r(f.imposto))
                    _par(
                        "IR fonte day trade: mês / anterior / a compensar",
                        f"{_r(f.irrf_day_trade_mes)} / {_r(f.irrf_day_trade_anterior)} / "
                        f"{_r(f.irrf_day_trade_a_compensar)}",
                    )
                    _par(
                        "IR fonte (Lei 11.033): mês / anterior / a compensar",
                        f"{_r(f.irrf_mes)} / {_r(f.irrf_anterior)} / {_r(f.irrf_a_compensar)}",
                    )
                    _par("Imposto a pagar / pago", f"{_r(f.a_pagar)} / {_r(f.pago)}")
    ui.label(f"Movimentos do mês com resultado ({len(movs)})").classes("coir-campo-rotulo")
    if not movs:
        ui.label("Nenhuma venda ou resultado neste mês.").classes("coir-vazio w-full")
        return
    valores: list[tuple[Valor, ...]] = [
        (
            x.data,
            x.ticker,
            CURTA[x.categoria],
            x.regra,
            x.cv or "",
            x.quantidade,
            x.valor,
            x.custo_baixa,
            x.resultado,
            # units fora do limite (ex.: UBBR11 em 2008): tributadas, nunca marcadas "isento"
            "isento"
            if conta_na_isencao(x, params) and comum and comum.isento
            else ("PME" if x.isento_pme else ""),
        )
        for x in movs
    ]
    tabela.grade(
        COLUNAS_MOV,
        valores,
        chaves=[carteira.nota_de(x.origem, mapa) for x in movs],
        altura="h-[420px]",
        pelo_conteudo=True,
        ao_clicar=lambda n: n is not None and ui.navigate.to(f"/notas/{n}"),
    ).mark("movimentos-mes")
    ui.label("Clique num movimento para abrir a nota de origem.").classes("coir-subtitulo text-xs")
