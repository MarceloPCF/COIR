"""A tela Eventos montada: as abas, a de conferir e o lançamento automático ao abrir."""

from collections.abc import Callable

from nicegui import run, ui

from coir.erros import CoirError
from coir.servicos import eventos
from coir.servicos.estado import Estado
from coir.servicos.eventos import EventoLancado
from coir.ui import ocupado, tabela
from coir.ui.telas.eventos.abas import ABAS, Aba, _Tela
from coir.ui.telas.eventos.conferir import _Conferir
from coir.ui.telas.eventos.painel import _Painel


class _AbaEventos:
    """Uma aba de tarefa: à esquerda os eventos dela, à direita o formulário."""

    def __init__(self, tela: _Tela, aba: Aba) -> None:
        self.aba = aba
        ui.label(aba.explicacao).classes("coir-subtitulo")
        with ui.row().classes("w-full items-start gap-4 flex-nowrap"):
            self.lista = ui.column().classes("min-w-[560px] flex-[1.4] gap-2")
            self.painel = _Painel(tela, aba.chave, aba)
        self.painel.desenhar()

    def mostrar(self, lancados: list[EventoLancado]) -> None:
        aba = self.aba
        meus = [e for e in lancados if e.dados.tipo in aba.tipos]
        self.lista.clear()
        with self.lista:
            ui.label(f"{aba.titulo} ({len(meus)})").classes("coir-titulo text-lg")
            if not meus:
                ui.label(aba.vazio).classes("coir-vazio w-full")
                return
            tabela.grade(
                aba.colunas, [aba.linha(e) for e in meus],
                chaves=list(range(len(meus))), altura="h-[520px]",
                ao_clicar=lambda i: self.painel.abrir_evento(meus[int(str(i))]),
            ).mark(f"grade-{aba.chave}")  # fmt: skip


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Eventos").classes("coir-titulo-pagina")
    ui.label(
        "O que muda a sua posição sem passar por uma nota de corretagem. Cada evento é conferido "
        "com as suas operações antes de ser gravado."
    ).classes("coir-subtitulo")
    if aviso_b3 := estado.sem_fonte_b3("eventos"):
        ui.label(aviso_b3).classes("coir-carimbo text-sm").mark("sem-fonte-b3")
    if not estado.existe or estado.titular is None:
        ui.label("Cadastre o titular primeiro.").classes("coir-vazio w-full")
        return
    titular = estado.titular
    contas = eventos.contas_do_titular(estado.banco, titular)
    if not contas:
        ui.label("Importe as notas primeiro: os eventos ficam numa conta com notas.").classes(
            "coir-vazio w-full"
        )
        return

    async def recarregar() -> None:
        try:
            lancados = await run.io_bound(eventos.listar, estado.banco, titular)
            propostas = await run.io_bound(estado.propostas)
            lacuna = await run.io_bound(eventos.lacuna_movimentacao, estado.banco, titular)
        except CoirError as erro:
            if conferir.lista.is_deleted:
                return
            conferir.lista.clear()
            with conferir.lista:
                ui.label(f"Não foi possível ler os eventos: {erro}").classes("coir-negativo")
            return
        if conferir.lista.is_deleted:
            return  # a tela foi refeita (titular, ano) ou fechada durante a leitura
        conferir.mostrar(propostas or [], lacuna)
        conferir.mostrar_sem_custo(lancados or [])
        guia_conferir.props(f'label="A conferir ({len(conferir.pendentes)})"')
        for aba in tarefas:
            aba.mostrar(lancados or [])

    tela = _Tela(estado, titular, contas, ao_mudar, recarregar)
    with ui.tabs().props("no-caps align=left").classes("w-full") as guias:
        guia_conferir = ui.tab("conferir", "A conferir").mark("aba-conferir")
        guia = {a.chave: ui.tab(a.chave, a.rotulo).mark(f"aba-{a.chave}") for a in ABAS}
    tarefas: list[_AbaEventos] = []
    with ui.tab_panels(guias, value=guia_conferir, animated=False).classes("w-full bg-transparent"):
        with ui.tab_panel(guia_conferir).classes("p-0 pt-3 gap-3"):
            conferir = _Conferir(tela)
        for a in ABAS:
            with ui.tab_panel(guia[a.chave]).classes("p-0 pt-3 gap-3"):
                tarefas.append(_AbaEventos(tela, a))

    def corrigir(e: EventoLancado) -> None:
        """Abre a bonificação sem custo no formulário da aba Lançados."""
        guias.set_value(guia["lancados"])
        next(x for x in tarefas if x.aba.chave == "lancados").painel.abrir_evento(e)

    conferir.corrigir = corrigir

    async def abrir() -> None:
        """Ao abrir a tela, lança o que não pede nada (notas importadas, tabelas novas)."""
        try:
            with ocupado.processando("Conferindo os eventos das suas notas"):
                automaticos = await run.io_bound(estado.lancar_eventos_automaticos)
        except CoirError as erro:
            ui.notify(f"Eventos não lançados: {erro}", type="warning", multi_line=True)
        else:
            if automaticos and (automaticos.ids or automaticos.recusadas):
                conferir.lote = (*conferir.lote, *automaticos.ids)
                conferir.recusadas = automaticos.recusadas
                if automaticos.ids:
                    ao_mudar()
        await recarregar()

    ui.timer(0.01, abrir, once=True)
