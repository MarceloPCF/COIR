"""Tela 11 — Outros custos dedutíveis (ex.: taxas de aluguel de ações - BTC): a grade do ano, mês
por categoria, editável; cada valor reduz o resultado da categoria no mês."""

from collections.abc import Callable
from decimal import Decimal

from nicegui import run, ui

from coir.domain.fiscal import Categoria
from coir.erros import CoirError
from coir.reports.fiscais import MESES
from coir.reports.modelo import TipoColuna, formatar
from coir.servicos import outros_custos
from coir.servicos.estado import Estado
from coir.ui.entrada import valor_br

ZERO = Decimal(0)
NOMES = {Categoria.COMUM: "Comum", Categoria.DAY_TRADE: "Day trade"}


def _r(v: Decimal) -> str:
    return formatar(v, TipoColuna.DINHEIRO)


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Outros custos dedutíveis").classes("coir-titulo-pagina")
    ui.label(
        f"Custos de {estado.ano} que não estão nas notas e reduzem o resultado do mês, como as "
        "taxas de aluguel de ações (BTC) cobradas pela corretora. Digite o total do mês em cada "
        "categoria; zero apaga. A apuração e a DARF são refeitas."
    ).classes("coir-subtitulo")
    if not estado.existe or estado.titular is None or estado.ano is None:
        ui.label("Cadastre o titular primeiro.").classes("coir-vazio w-full")
        return
    titular, ano = estado.titular, estado.ano
    atuais = {
        (c.mes, c.categoria): c for c in outros_custos.listar(estado.banco, titular) if c.ano == ano
    }
    total = ui.label("").classes("coir-numero")

    def somar() -> None:
        soma = sum((c.valor for c in atuais.values()), ZERO)
        total.set_text(f"Total do ano: {_r(soma)}")

    async def gravar(mes: int, categoria: Categoria, texto: str, descricao: str) -> None:
        valor = valor_br(texto or "0")
        if valor is None:
            ui.notify("Valor inválido.", type="warning")
            return
        try:
            await run.io_bound(
                outros_custos.gravar, estado.banco, titular, ano, mes, categoria, valor, descricao
            )
        except CoirError as erro:
            ui.notify(str(erro), type="negative")
            return
        estado.invalidar()
        ao_mudar()
        lista = await run.io_bound(outros_custos.listar, estado.banco, titular)
        atuais.clear()
        atuais.update({(c.mes, c.categoria): c for c in lista or [] if c.ano == ano})
        somar()
        if valor:
            ui.notify(
                f"{MESES[mes - 1]}: reduz o resultado {NOMES[categoria].lower()} em {_r(valor)}.",
                type="positive",
            )

    with (
        ui.grid(columns="120px 160px 160px 1fr")
        .classes("coir-folha w-full gap-x-4 gap-y-1 items-center")
        .mark("grade-outros")
    ):
        for rotulo in ("Mês", "Comum", "Day trade", "Descrição"):
            ui.label(rotulo).classes("coir-campo-rotulo")
        for mes in range(1, 13):
            ui.label(MESES[mes - 1])
            campos = {}
            for categoria in (Categoria.COMUM, Categoria.DAY_TRADE):
                atual = atuais.get((mes, categoria))
                campos[categoria] = (
                    ui.input(value=_r(atual.valor) if atual else "")
                    .props("outlined dense input-class=text-right")
                    .mark(f"custo-{mes}-{categoria.value}")
                )
            descricao_atual = next(
                (c.descricao for c in (atuais.get((mes, k)) for k in NOMES) if c and c.descricao),
                "",
            )
            descricao = (
                ui.input(value=descricao_atual, placeholder="ex.: aluguel de ações")
                .props("outlined dense")
                .mark(f"descricao-{mes}")
            )
            for categoria, campo in campos.items():
                campo.on("blur", _salvar(gravar, mes, categoria, campo, descricao))
    somar()


def _salvar(
    gravar: Callable[[int, Categoria, str, str], object],
    mes: int,
    categoria: Categoria,
    campo: ui.input,
    descricao: ui.input,
) -> Callable[[], object]:
    return lambda: gravar(mes, categoria, campo.value or "", descricao.value or "")
