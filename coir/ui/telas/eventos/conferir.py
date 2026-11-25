"""A aba Conferir da tela Eventos: propostas da B3 e do pacote público, eventos seguros e ativos
que pararam de negociar."""

from collections.abc import Callable

from nicegui import run, ui

from coir.erros import CoirError
from coir.servicos import eventos
from coir.servicos.eventos import EventoLancado, Proposta
from coir.ui import ocupado, tabela
from coir.ui.telas.eventos.abas import _Tela
from coir.ui.telas.eventos.painel import _Painel
from coir.ui.telas.eventos.textos import COLUNAS_PROPOSTAS, COLUNAS_SEGURAS, TIPOS


class _Conferir:
    """Aba "A conferir": os avisos de ativos que saíram da bolsa, as propostas da B3, o botão das
    seguras e o desfazer do último lote."""

    def __init__(self, tela: _Tela) -> None:
        self.tela = tela
        self.ver_todas = False
        self.lote: tuple[int, ...] = ()
        self.recusadas: tuple[tuple[str, str], ...] = ()
        self.propostas: list[Proposta] = []
        ui.label(
            "Desdobramentos, grupamentos e bonificações com custo publicado entram sozinhos para "
            "as ações que estavam em carteira. Aqui fica o que depende de você, como subscrição "
            "e bonificação sem custo."
        ).classes("coir-subtitulo")
        self.lacuna = (
            ui.label("").classes("coir-carimbo text-sm w-full").mark("lacuna-movimentacao")
        )
        self.lacuna.set_visibility(False)
        self.corrigir: Callable[[EventoLancado], None] | None = None  # abre o evento em Lançados
        self.secao_sem_custo = ui.column().classes("w-full gap-1")
        self.secao_sem_custo.set_visibility(False)
        self.secao_saidas = ui.column().classes("w-full gap-1")
        self.secao_lote = ui.column().classes("w-full gap-1")
        self.secao_lote.set_visibility(False)
        with ui.row().classes("w-full items-start gap-4 flex-nowrap"):
            self.lista = ui.column().classes("min-w-[560px] flex-[1.4] gap-2")
            self.painel = _Painel(tela, "conferir", None)
        self.painel.desenhar()
        self._desenhar_saidas()

    @property
    def pendentes(self) -> list[Proposta]:
        return [p for p in self.propostas
                if p.situacao == "pendente" and not p.depois_da_ultima_nota]  # fmt: skip

    def mostrar(self, propostas: list[Proposta], lacuna: str | None = None) -> None:
        self.lacuna.set_text(lacuna or "")
        self.lacuna.set_visibility(bool(lacuna))
        self.propostas = propostas
        self._desenhar_lista()
        self._desenhar_lote()
        aberta = self.painel.form.proposta
        if aberta is not None:  # outra aba gravou: a quantidade pode ter mudado
            atual = next((p for p in propostas if p.chave == aberta.chave), None)
            if atual is None:
                self.painel.novo()
            else:
                self.painel.abrir_proposta(atual)

    def mostrar_sem_custo(self, lancados: list[EventoLancado]) -> None:
        """Bonificações já gravadas sem custo atribuído: o preço médio da venda sai errado."""
        sem_custo = eventos.bonificacoes_sem_custo_de(lancados)
        self.secao_sem_custo.clear()
        self.secao_sem_custo.set_visibility(bool(sem_custo))
        with self.secao_sem_custo:
            for e in sem_custo:
                with ui.row().classes("w-full items-center gap-2 flex-nowrap"):
                    ui.label(eventos.aviso_sem_custo(e)).classes(
                        "coir-carimbo text-sm grow").mark("bonificacao-sem-custo")  # fmt: skip
                    if self.corrigir is not None:
                        ui.button("Corrigir", on_click=lambda e=e, ir=self.corrigir: ir(e)
                                  ).props("flat dense").mark("corrigir-sem-custo")  # fmt: skip

    def _desenhar_saidas(self) -> None:
        t = self.tela
        try:
            saidas = t.estado.ativos_que_sairam()
        except CoirError as erro:  # ex.: ticker sem tipo de ativo; o resto da tela abre
            self.secao_saidas.clear()
            self.secao_saidas.set_visibility(True)
            with self.secao_saidas:
                ui.label(f"Não foi possível conferir os ativos que saíram da bolsa: {erro}"
                         ).classes("coir-negativo text-sm")  # fmt: skip
            return
        self.secao_saidas.clear()
        self.secao_saidas.set_visibility(bool(saidas))
        with self.secao_saidas:
            for saiu in saidas:
                with ui.row().classes("w-full items-center gap-2 flex-nowrap"):
                    ui.label(saiu.texto).classes("coir-subtitulo text-sm grow").mark(
                        "ativo-que-saiu"
                    )

                    async def dispensar(chave: str = saiu.chave) -> None:
                        await run.io_bound(eventos.ignorar, t.estado.banco, t.titular, chave,
                                           t.estado.agora())  # fmt: skip
                        t.ao_mudar()  # o sino do topo conta o aviso
                        if not self.secao_saidas.is_deleted:
                            self._desenhar_saidas()

                    ui.button("Ignorar", on_click=dispensar).props("flat dense").mark(
                        "ignorar-saida"
                    )

    def _desenhar_lista(self) -> None:
        propostas, pendentes = self.propostas, self.pendentes
        mostradas = propostas if self.ver_todas else pendentes
        seguras = [p for p in pendentes if p.segura]
        posteriores = [p for p in propostas
                       if p.situacao == "pendente" and p.depois_da_ultima_nota]  # fmt: skip
        self.lista.clear()
        with self.lista:
            with ui.row().classes("w-full items-center gap-3"):
                if seguras:
                    ui.button(f"Lançar as seguras ({len(seguras)})", icon="done_all",
                              on_click=lambda: self._confirmar(seguras)).props("unelevated").mark(
                        "lancar-seguras").tooltip(
                        "Desdobramentos e grupamentos das tabelas da B3, e bonificações com o "
                        "custo publicado pela companhia na CVM.")  # fmt: skip
                ui.space()
                ui.checkbox("mostrar lançadas, ignoradas e posteriores", value=self.ver_todas,
                            on_change=lambda e: self._ver_todas(bool(e.value))
                            ).mark("ver-propostas")  # fmt: skip
            if posteriores and not self.ver_todas:
                ui.label(f"{len(posteriores)} evento(s) depois da sua última nota ficaram de fora: "
                         "o COIR não sabe se você ainda tinha o ativo. Importe as notas seguintes, "
                         "ou marque a opção acima para ver.").classes(
                    "coir-subtitulo text-sm w-full").mark("eventos-posteriores")  # fmt: skip
            if not propostas:
                ui.label("Nada a conferir: nenhum evento público da B3 atinge as suas posições. A "
                         "movimentação da B3 (Importar > Planilhas da B3) pode trazer mais "
                         "sugestões.").classes(
                    "coir-vazio w-full")  # fmt: skip
            elif not mostradas:
                ui.label("Tudo conferido: os eventos sugeridos já estão lançados ou "
                         "ignorados.").classes("coir-vazio w-full")  # fmt: skip
            else:
                tabela.grade(
                    COLUNAS_PROPOSTAS,
                    [(p.data, TIPOS[p.dados.tipo], p.codigo, p.corretora or "", p.quantidade,
                      p.valor, "depois da última nota" if p.situacao == "pendente"
                      and p.depois_da_ultima_nota else p.situacao) for p in mostradas],
                    chaves=list(range(len(mostradas))), altura="h-[420px]",
                    ao_clicar=lambda i: self.painel.abrir_proposta(mostradas[int(str(i))]),
                ).mark("grade-propostas")  # fmt: skip

    def _ver_todas(self, sim: bool) -> None:
        self.ver_todas = sim
        self._desenhar_lista()

    def _confirmar(self, seguras: list[Proposta]) -> None:
        n = len(seguras)
        plural = "s" if n > 1 else ""
        with ui.dialog() as janela, ui.card().classes("w-[900px] max-w-full gap-3"):
            ui.label(f"Lançar {n} evento{plural}?").classes("coir-titulo text-base")
            ui.label(
                "A carteira e a apuração serão refeitas. Se algo não fechar com as suas "
                "operações, o evento não entra e o motivo aparece na tela. Depois, dá para "
                "desfazer."
            ).classes("text-sm")
            tabela.grade(
                COLUNAS_SEGURAS,
                [(p.data, TIPOS[p.dados.tipo], p.codigo, p.corretora or "", p.dados.quantidade,
                  p.dados.valor or None, p.custo_por_acao,
                  p.referencia if p.custo_por_acao else "") for p in seguras],
                altura="h-[300px]",
            ).mark("grade-seguras")  # fmt: skip

            async def confirmar() -> None:
                janela.close()
                await self._lancar(seguras)

            with ui.row().classes("w-full justify-end gap-2"):
                ui.button("Agora não", on_click=janela.close).props("flat").mark("seguras-nao")
                ui.button(f"Lançar {n} evento{plural}", icon="done_all",
                          on_click=confirmar).props("unelevated").mark(
                    "seguras-sim")  # fmt: skip
        janela.open()

    async def _lancar(self, seguras: list[Proposta]) -> None:
        t = self.tela
        estado = t.estado
        try:
            with ocupado.processando("Lançando os eventos"):
                resultado = await run.io_bound(
                    eventos.lancar_seguras, estado.banco, t.titular, [p.chave for p in seguras],
                    estado.refs, estado.fiscais, estado.pasta_refs,
                )  # fmt: skip
        except CoirError as erro:
            ui.notify(str(erro), type="negative", multi_line=True, timeout=12000)
            return
        if resultado is None:
            return
        self.lote, self.recusadas = (*self.lote, *resultado.ids), resultado.recusadas
        estado.invalidar()
        t.ao_mudar()
        if self.lista.is_deleted:
            return  # a tela foi refeita (titular, ano) ou fechada durante o lançamento
        self.painel.novo()
        await t.recarregar()

    def _desenhar_lote(self) -> None:
        self.secao_lote.clear()
        self.secao_lote.set_visibility(bool(self.lote or self.recusadas))
        with self.secao_lote:
            if self.lote:
                with ui.row().classes("w-full items-center gap-3"):
                    ui.label(f"{len(self.lote)} evento(s) lançado(s) sem precisar de conferência. "
                             "Carteira e apuração refeitas.").classes(
                        "coir-positivo").mark("lote-lancado")  # fmt: skip
                    ui.button("Desfazer", icon="undo", on_click=self._desfazer).props(
                        "flat dense").mark("desfazer-lote")  # fmt: skip
            for ativo, motivo in self.recusadas:
                ui.label(f"{ativo} não entrou: {motivo}").classes("coir-negativo text-sm")

    async def _desfazer(self) -> None:
        t = self.tela
        try:
            await run.io_bound(eventos.excluir_lote, t.estado.banco, t.titular, list(self.lote),
                               t.estado.refs, t.estado.fiscais, True,
                               t.estado.pasta_refs)  # fmt: skip
        except CoirError as erro:
            ui.notify(f"Não foi possível desfazer: {erro}", type="negative", multi_line=True,
                      timeout=12000)  # fmt: skip
            return
        self.lote, self.recusadas = (), ()
        t.estado.invalidar()
        t.ao_mudar()
        ui.notify("Lançamento desfeito.", type="positive")
        if self.lista.is_deleted:
            return
        self.painel.novo()
        await t.recarregar()
