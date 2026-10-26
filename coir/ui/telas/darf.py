"""Tela 13 — DARF: as DARFs do ano com a situação e, ao escolher uma, o documento em campos
numerados (como no formulário), a multa de mora e o registro do pagamento com o comprovante."""

import calendar
from collections.abc import Callable
from datetime import date
from decimal import Decimal

from nicegui import events, run, ui

from coir.erros import CoirError
from coir.reports.fiscais import MESES
from coir.reports.modelo import Coluna, TipoColuna, Valor, formatar
from coir.servicos import darf
from coir.servicos.estado import Estado
from coir.ui import tabela
from coir.ui.entrada import valor_br

T = TipoColuna
COLUNAS = (
    Coluna("Competência"),
    Coluna("Vencimento", T.DATA),
    Coluna("Principal", T.DINHEIRO),
    Coluna("Multa", T.DINHEIRO),
    Coluna("Total", T.DINHEIRO),
    Coluna("Situação"),
    Coluna("Pago em", T.DATA),
)
CLASSES = {
    darf.Situacao.EM_ATRASO: "coir-negativo",
    darf.Situacao.A_PAGAR: "coir-carimbo",
    darf.Situacao.PAGA: "coir-positivo",
    darf.Situacao.ACUMULADA: "coir-subtitulo",
    darf.Situacao.EM_APURACAO: "coir-subtitulo",
}


def _r(valor: Decimal) -> str:
    return formatar(valor, T.DINHEIRO)


def _valores(linhas: list[darf.LinhaDarf]) -> list[tuple[Valor, ...]]:
    return [
        (
            f"{x.mes:02d}/{x.ano}",
            x.vencimento,
            x.principal,
            x.multa.valor if x.multa else None,
            x.total,
            x.situacao.value,
            x.pagamento.data if x.pagamento else None,
        )
        for x in linhas
    ]


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("DARF").classes("coir-titulo-pagina")
    ui.label(
        f"As DARFs de {estado.ano} de renda variável, o vencimento, o pagamento, a multa "
        "e os juros de mora por atraso (Lei 9.430/1996 art. 61)."
    ).classes("coir-subtitulo")
    with ui.row().classes("w-full items-start gap-4"):
        lista = ui.column().classes("min-w-[520px] flex-[1.2] gap-2")
        lado = ui.column().classes("min-w-[380px] flex-1 gap-2")
    linhas: list[darf.LinhaDarf] = []

    async def carregar(escolher: tuple[int, int] | None = None) -> None:
        lista.clear()
        with lista:
            ui.spinner(size="lg")
        try:
            resultado = await run.io_bound(estado.darfs)
        except CoirError as erro:
            lista.clear()
            with lista:
                ui.label(f"Não foi possível apurar: {erro}").classes("coir-negativo")
            return
        if resultado is None:
            return
        linhas[:] = resultado
        lista.clear()
        with lista:
            if not linhas:
                ui.label("Nenhuma DARF neste ano.").classes("coir-vazio w-full")
            else:
                tabela.grade(
                    COLUNAS,
                    _valores(linhas),
                    chaves=list(range(len(linhas))),
                    ao_clicar=lambda i: abrir(int(str(i))),
                ).mark("grade-darf")
                atrasadas = [x for x in linhas if x.situacao is darf.Situacao.EM_ATRASO]
                if atrasadas:
                    ui.label(
                        f"{len(atrasadas)} DARF(s) em atraso: a multa mostrada é estimada até "
                        f"hoje ({estado.agora():%d/%m/%Y})."
                    ).classes("coir-negativo text-sm")
        lado.clear()
        indice = next((i for i, x in enumerate(linhas) if (x.ano, x.mes) == escolher), None)
        if indice is not None:
            abrir(indice)
        else:
            with lado:
                ui.label("Escolha uma DARF na lista.").classes("coir-vazio w-full")

    def abrir(indice: int) -> None:
        lado.clear()
        with lado:
            _documento(estado, linhas[indice], recarregar)

    async def recarregar(competencia: tuple[int, int]) -> None:
        ao_mudar()
        await carregar(competencia)

    escolhida, estado.darf_escolhida = estado.darf_escolhida, None
    ui.timer(0.01, lambda: carregar(escolhida), once=True)


def _campo(numero: str, rotulo: str, valor: str, classe: str = "") -> None:
    with ui.row().classes("coir-item-lista w-full items-baseline gap-3 py-1 no-wrap"):
        ui.label(numero).classes("coir-campo-rotulo w-6")
        ui.label(rotulo).classes("text-sm grow")
        ui.label(valor).classes(f"coir-numero text-right {classe}")


def _documento(
    estado: Estado, x: darf.LinhaDarf, recarregar: Callable[[tuple[int, int]], object]
) -> None:
    periodo = date(x.ano, x.mes, calendar.monthrange(x.ano, x.mes)[1])
    with ui.row().classes("items-baseline gap-3"):
        ui.label(f"DARF de {MESES[x.mes - 1].lower()}/{x.ano}").classes("coir-titulo text-lg")
        ui.label(x.situacao.value).classes(f"text-sm font-semibold {CLASSES[x.situacao]}")
    if x.situacao is darf.Situacao.EM_APURACAO:
        ui.label(
            f"O mês ainda não acabou: {_r(x.principal)} é o imposto apurado até hoje e muda com as "
            "próximas notas. A DARF sai depois do fim do mês."
        ).classes("coir-vazio w-full")
        return
    if x.situacao is darf.Situacao.ACUMULADA:
        ui.label(
            f"{_r(x.principal)} abaixo do valor mínimo da DARF: soma-se à DARF do mês seguinte "
            "e vence com ela, sem multa (Lei 9.430/1996 art. 68 §1º)."
        ).classes("coir-vazio w-full")
        return
    with ui.column().classes("coir-folha w-full gap-0"):
        _campo("02", "Período de apuração", f"{periodo:%d/%m/%Y}")
        _campo("04", "Código da receita", x.codigo)
        _campo("06", "Data de vencimento", f"{x.vencimento:%d/%m/%Y}" if x.vencimento else "")
        _campo("07", "Valor do principal", _r(x.principal))
        multa = x.multa
        _campo(
            "08",
            "Valor da multa",
            _r(multa.valor) if multa else "0,00",
            "coir-negativo" if multa and multa.valor else "",
        )
        juros = x.juros
        _campo("09", "Valor dos juros", _r(juros.valor) if juros else "0,00",
               "coir-negativo" if juros and juros.valor else "")  # fmt: skip
        _campo("10", "Valor total", _r(x.total), "font-semibold")
    if multa and multa.dias:
        ate = "pagamento" if x.pagamento else "hoje (estimativa)"
        ajuste = " (data não útil: vale o próximo dia útil)" if multa.ajustado else ""
        ui.label(
            f"Multa: {multa.dias} dia(s) de atraso até o {ate}, em {multa.pagamento:%d/%m/%Y}"
            f"{ajuste}: {formatar(multa.percentual, T.PERCENTUAL)} do principal (com o teto)."
        ).classes("coir-subtitulo text-xs")
        ui.label(
            "A contagem segue a Receita (Sicalc): do primeiro dia útil após o vencimento. Pode "
            "haver discussão sobre a exclusão da multa (CTN art. 138; STJ REsp 1.149.022/SP; "
            "Súmula 360): consulte um profissional."
        ).classes("coir-subtitulo text-xs")
    juros = x.juros
    if juros and juros.valor:
        selic = " + ".join(f"{m} {formatar(t, T.QUANTIDADE)}%" for m, t in juros.taxas)
        ui.label(
            f"Juros: {formatar(juros.percentual, T.PERCENTUAL)} do principal = Selic dos meses "
            f"entre o vencimento e o pagamento ({selic or 'nenhum mês'}) + 1% no mês do "
            f"pagamento ({juros.pagamento:%m/%Y})."
        ).classes("coir-subtitulo text-xs")
    if juros and juros.meses_sem_taxa:
        faltam = ", ".join(juros.meses_sem_taxa)
        texto = (f"A tabela Selic do COIR não tem {faltam}: atualize a tabela; os juros estão "
                 "menores que os devidos." if x.pagamento else
                 f"Estimativa: a Selic de {faltam} ainda não está na tabela; os juros finais "
                 "serão maiores. Confira no Sicalc antes de pagar.")  # fmt: skip
        ui.label(texto).classes("coir-carimbo text-xs")
    p = x.pagamento
    if p is not None:
        with ui.column().classes("coir-folha w-full gap-1"):
            ui.label("Pagamento registrado").classes("coir-titulo text-base")
            ui.label(f"Pago em {p.data:%d/%m/%Y}: {_r(p.valor)}").classes("coir-numero")
            dif = x.diferenca
            if dif is not None and dif < 0:
                ui.label(
                    f"Pago a menor: faltam {_r(-dif)} (gere a DARF complementar no Sicalc)."
                ).classes("coir-negativo text-sm")
            elif dif is not None and dif > 0 and x.juros and x.juros.meses_sem_taxa:
                ui.label(
                    f"Diferença de {_r(dif)}: provavelmente a Selic dos meses ainda sem taxa na "
                    "tabela do COIR."
                ).classes("coir-subtitulo text-sm")
            elif dif is not None and dif > 0:
                ui.label(
                    f"O valor registrado é {_r(dif)} maior que o devido para essa data. Se foi "
                    "engano no registro, corrija abaixo; se pagou mesmo a mais, a diferença pode "
                    "ser pedida de volta à Receita (PER/DCOMP)."
                ).classes("coir-carimbo text-sm")
            if p.comprovante:
                ui.label(f"Comprovante: {p.comprovante}").classes("coir-subtitulo text-xs")
            if p.observacao:
                ui.label(p.observacao).classes("coir-subtitulo text-sm")

            async def desfazer() -> None:
                await run.io_bound(
                    darf.excluir_pagamento, estado.banco, estado.titular or 0, x.ano, x.mes
                )
                estado.invalidar()
                ui.notify("Registro de pagamento desfeito.")
                await _chamar(recarregar, (x.ano, x.mes))

            ui.button("Desfazer o registro", icon="undo", on_click=desfazer).props(
                "flat dense"
            ).mark("desfazer-pagamento")
    _formulario(estado, x, recarregar)


async def _chamar(f: Callable[[tuple[int, int]], object], arg: tuple[int, int]) -> None:
    resultado = f(arg)
    if hasattr(resultado, "__await__"):
        await resultado  # type: ignore[misc]


def _formulario(
    estado: Estado, x: darf.LinhaDarf, recarregar: Callable[[tuple[int, int]], object]
) -> None:
    hoje = estado.agora().date()
    sugestao = x.multa.pagamento if x.multa else hoje
    comprovante: list[tuple[str, bytes]] = []
    if x.pagamento:
        caixa = ui.expansion("Alterar o pagamento", icon="edit", value=False)
    else:  # aberto e fixo: o título não fecha o formulário
        caixa = ui.column().classes("gap-2")
        with caixa:
            ui.label("Registrar o pagamento").classes("coir-titulo text-base")
            ui.label(
                "Informe a data e o valor pagos (com multa e juros, se pagou em atraso) e salve: a "
                "DARF passa a constar como paga."
            ).classes("coir-subtitulo text-xs")
    with caixa.classes("coir-folha w-full").mark("form-pagamento"):
        data = (
            ui.input(
                "Data do pagamento",
                value=(x.pagamento.data if x.pagamento else sugestao).isoformat(),
            )
            .props("outlined dense type=date")
            .classes("w-[200px]")
            .mark("data-pagamento")
        )
        valor = (
            ui.input("Valor pago (R$)", value=_r(x.pagamento.valor if x.pagamento else x.total))
            .props("outlined dense")
            .classes("w-[200px]")
            .mark("valor-pagamento")
        )
        observacao = ui.input("Observação").props("outlined dense").classes("w-full")
        sugerido = {"valor": valor.value}

        def nova_data() -> None:
            """Troca a data: o valor sugerido acompanha (multa e juros até essa data), a não ser
            que o valor já tenha sido digitado."""
            try:
                dia = date.fromisoformat(data.value or "")
            except ValueError:
                return
            if valor.value == sugerido["valor"]:
                sugerido["valor"] = _r(darf.devido_em(x, dia, estado.fiscais))
                valor.set_value(sugerido["valor"])

        data.on_value_change(nova_data)

        async def receber(e: events.UploadEventArguments) -> None:
            comprovante[:] = [(e.file.name, await e.file.read())]
            ui.notify(f"Comprovante: {e.file.name}")

        ui.upload(
            label="Comprovante (PDF, opcional)", auto_upload=True, on_upload=receber, max_files=1
        ).props('accept=".pdf,application/pdf" flat').classes("coir-envio w-full")

        async def salvar() -> None:
            try:
                dia = date.fromisoformat(data.value or "")
            except ValueError:
                ui.notify("Data inválida.", type="warning")
                return
            pago = valor_br(valor.value or "")
            if pago is None:
                ui.notify("Valor inválido.", type="warning")
                return
            try:
                await run.io_bound(
                    darf.registrar_pagamento,
                    estado.banco,
                    estado.dados,
                    estado.titular or 0,
                    x.ano,
                    x.mes,
                    dia,
                    pago,
                    estado.agora(),
                    comprovante[0] if comprovante else None,
                    observacao.value or "",
                )
            except CoirError as erro:
                ui.notify(str(erro), type="negative")
                return
            estado.invalidar()
            ui.notify("Pagamento registrado.", type="positive")
            await _chamar(recarregar, (x.ano, x.mes))

        ui.button("Salvar o pagamento", icon="task_alt", on_click=salvar).mark("salvar-pagamento")
