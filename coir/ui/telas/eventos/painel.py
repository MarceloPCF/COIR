"""O formulário de lançamento de cada aba da tela Eventos: campos por tipo, gravar, excluir e
ignorar."""

from dataclasses import replace
from datetime import date
from decimal import Decimal

from nicegui import run, ui

from coir.domain.fiscal import NaturezaEvento, TipoEvento
from coir.erros import CoirError
from coir.servicos import eventos
from coir.servicos.eventos import AVISO_INCORPORACAO_ACOES, DadosEvento, EventoLancado, Proposta
from coir.ui.entrada import valor_br
from coir.ui.telas.eventos.abas import Aba, _Form, _numero, _Tela
from coir.ui.telas.eventos.textos import (
    AJUDA,
    COM_DESTINO,
    COM_OUTRA_CLASSE,
    COM_RECEBIDO,
    NATUREZAS,
    NOS_LANCADOS,
    ROTULO_ATIVO,
    ROTULO_DATA,
    ROTULO_QUANTIDADE,
    ROTULO_RECEBIDO,
    ROTULO_VALOR,
    SEM_CUSTO,
    SEM_QUANTIDADE,
    SENTIDO,
    TIPOS,
)


class _Painel:
    """O formulário de uma aba; sem ``aba``, o da aba "A conferir", que só abre uma proposta."""

    def __init__(self, tela: _Tela, chave: str, aba: Aba | None) -> None:
        self.tela, self.chave, self.aba = tela, chave, aba
        self.form = _Form(self._vazio())
        self.area = ui.column().classes("coir-folha min-w-[380px] flex-1 gap-2").mark(
            self.m("painel-evento"))  # fmt: skip

    def m(self, nome: str) -> str:
        return f"{self.chave}-{nome}"

    def _vazio(self) -> DadosEvento | None:
        if self.aba is None:
            return None
        tipo = self.aba.novos[0]
        natureza = NaturezaEvento.CONVERSAO if tipo is TipoEvento.INCORPORACAO else None
        return DadosEvento(tipo=tipo, data=self.tela.estado.fim_do_ano(),
                           conta=next(iter(self.tela.contas), None), ticker="",
                           cv=SENTIDO.get(tipo), natureza=natureza)  # fmt: skip

    def abrir_evento(self, e: EventoLancado) -> None:
        self.form = _Form(e.dados, e.id)
        self.desenhar()

    def abrir_proposta(self, p: Proposta) -> None:
        self.form = _Form(p.dados, proposta=p)
        self.desenhar()

    def novo(self) -> None:
        self.form = _Form(self._vazio())
        self.desenhar()

    def _ler(self) -> DadosEvento | None:
        """Os campos do painel -> DadosEvento (None se algum número está ilegível)."""
        c, d = self.form.campos, self.form.dados
        assert d is not None

        def texto(nome: str) -> str:
            return str(getattr(c[nome], "value", "") or "").strip() if nome in c else ""

        def numero(nome: str) -> Decimal | None:
            t = texto(nome)
            if not t:
                return None
            v = valor_br(t)
            if v is None:
                ui.notify(f"Número inválido: {t}", type="warning")
                raise ValueError(t)
            return v

        try:
            dia = date.fromisoformat(texto("data")) if texto("data") else d.data
        except ValueError:
            ui.notify(f"Data inválida: {texto('data')}", type="warning")
            return None
        try:
            percentual = numero("percentual")
            return replace(
                d,
                data=dia,
                conta=getattr(c["conta"], "value", d.conta),
                ticker=texto("ticker"),
                cv=texto("cv") or d.cv,
                quantidade=numero("quantidade"),
                valor=numero("valor") or Decimal(0),
                conta_destino=getattr(c["destino"], "value", None) if "destino" in c else None,
                ticker_destino=texto("ticker_destino") or None,
                quantidade_destino=numero("quantidade_destino"),
                natureza=(
                    NaturezaEvento(texto("natureza"))
                    if texto("natureza")
                    else d.natureza
                    if d.tipo in COM_OUTRA_CLASSE
                    else None
                ),
                percentual_custo=None if percentual is None else percentual / 100,
                valor_corrente=numero("valor_corrente"),
                observacao=texto("observacao"),
            )
        except ValueError:
            return None

    def _ler_outras(self) -> list[tuple[int | None, str, str]]:
        c = self.form.campos
        return [
            (getattr(c[f"conta_{i}"], "value", None),
             str(getattr(c[f"quantidade_{i}"], "value", "") or "").strip(),
             str(getattr(c[f"quantidade_destino_{i}"], "value", "") or "").strip())
            for i in range(len(self.form.outras))
        ]  # fmt: skip

    def _mais_uma_conta(self) -> None:
        self.form.dados = self._ler() or self.form.dados
        self.form.outras = [*self._ler_outras(), (None, "", "")]
        self.desenhar()

    def _lote(self, d: DadosEvento) -> list[tuple[DadosEvento, int | None]] | None:
        """O evento do painel e, na cisão, o das outras contas (mesmos dados, outra conta)."""
        saida: list[tuple[DadosEvento, int | None]] = [(d, self.form.evento_id)]
        for conta, qtd, recebida in self._ler_outras() if d.tipo is TipoEvento.CISAO else []:
            q, r = valor_br(qtd) if qtd else None, valor_br(recebida) if recebida else None
            if conta is None or q is None or r is None:
                ui.notify("Outra conta: informe a conta e as duas quantidades.", type="warning")
                return None
            saida.append((replace(d, conta=conta, quantidade=q, quantidade_destino=r), None))
        return saida

    def _mudar_tipo(self, valor: object) -> None:
        atual = self._ler() or self.form.dados
        assert atual is not None
        tipo = TipoEvento(str(valor))
        natureza = NaturezaEvento.CONVERSAO if tipo is TipoEvento.INCORPORACAO else None
        self.form.dados = replace(atual, tipo=tipo, cv=SENTIDO.get(tipo), natureza=natureza)
        self.desenhar()

    async def _gravar(self) -> None:
        d = self._ler()
        pedidos = None if d is None else self._lote(d)
        if pedidos is None:
            return
        t = self.tela
        try:
            await run.io_bound(eventos.salvar_lote, t.estado.banco, t.titular, pedidos,
                               t.estado.refs, t.estado.fiscais)  # fmt: skip
        except CoirError as erro:
            ui.notify(str(erro), type="negative", multi_line=True, timeout=12000)
            return
        t.estado.invalidar()
        t.ao_mudar()
        ui.notify("Evento gravado. Carteira e apuração refeitas.", type="positive")
        if self.area.is_deleted:
            return  # a tela foi refeita (titular, ano) ou fechada durante a gravação
        self.novo()
        await t.recarregar()

    async def _excluir(self) -> None:
        t = self.tela
        if self.form.evento_id is None:
            return
        try:
            await run.io_bound(eventos.excluir, t.estado.banco, t.titular, self.form.evento_id,
                               t.estado.refs, t.estado.fiscais, t.estado.pasta_refs)  # fmt: skip
        except CoirError as erro:
            ui.notify(str(erro), type="negative")
            return
        t.estado.invalidar()
        t.ao_mudar()
        ui.notify("Evento excluído.", type="positive")
        if self.area.is_deleted:
            return
        self.novo()
        await t.recarregar()

    async def _ignorar(self, sim: bool) -> None:
        t = self.tela
        if self.form.proposta is None:
            return
        await run.io_bound(eventos.ignorar, t.estado.banco, t.titular, self.form.proposta.chave,
                           t.estado.agora(), sim)  # fmt: skip
        t.ao_mudar()  # o sino do topo conta as propostas pendentes
        if self.area.is_deleted:
            return
        self.novo()
        await t.recarregar()

    def _opcoes_de_tipo(self, d: DadosEvento) -> dict[str, str]:
        tipos: frozenset[TipoEvento] | tuple[TipoEvento, ...]
        if self.aba is None:
            tipos = NOS_LANCADOS
        elif self.form.evento_id is not None and d.tipo not in self.aba.novos:
            tipos = (d.tipo,)
        else:
            tipos = self.aba.novos
        return {t.value: r for t, r in TIPOS.items() if t in tipos}

    def _titulo(self, d: DadosEvento) -> str:
        if self.form.proposta is not None:
            return f"Proposta da B3: {self.form.proposta.movimentacao} {self.form.proposta.codigo}"
        if self.form.evento_id is not None:
            return f"Corrigir: {TIPOS[d.tipo]}"
        return self.aba.titulo_novo if self.aba is not None else ""

    def desenhar(self) -> None:
        d, form, m = self.form.dados, self.form, self.m
        contas = self.tela.contas
        self.area.clear()
        form.campos = {}
        c = form.campos
        with self.area:
            if d is None:
                ui.label("Clique numa proposta para conferir e lançar, ou para ignorar.").classes(
                    "coir-subtitulo").mark(m("sem-proposta"))  # fmt: skip
                return
            ui.label(self._titulo(d)).classes("coir-titulo text-lg").mark(m("titulo-painel"))
            if form.proposta is not None:
                ui.label(form.proposta.nota).classes("coir-carimbo text-sm")
            opcoes = self._opcoes_de_tipo(d)
            if len(opcoes) > 1:
                ui.select(opcoes, value=d.tipo.value, label="O que houve",
                          on_change=lambda e: self._mudar_tipo(e.value)).props(
                    "outlined dense options-dense").classes("w-full").mark(
                    m("tipo-evento"))  # fmt: skip
            ui.label(AJUDA[d.tipo]).classes("coir-subtitulo text-xs")
            with ui.row().classes("w-full gap-2 flex-nowrap"):
                c["data"] = (
                    ui.input(ROTULO_DATA.get(d.tipo, "Data ex"), value=d.data.isoformat())
                    .props("outlined dense type=date")
                    .classes("flex-1")
                    .mark(m("data-evento"))
                )
                rotulo_conta = "Origem" if d.tipo is TipoEvento.TRANSFERENCIA else "Conta"
                c["conta"] = (
                    ui.select(contas, value=d.conta, label=rotulo_conta)
                    .props("outlined dense options-dense")
                    .classes("flex-1")
                    .mark(m("conta-evento"))
                )
            if d.tipo is TipoEvento.TRANSFERENCIA:
                c["destino"] = (
                    ui.select(contas, value=d.conta_destino, label="Destino")
                    .props("outlined dense options-dense")
                    .classes("w-full")
                    .mark(m("destino-evento"))
                )
            with ui.row().classes("w-full gap-2 flex-nowrap items-center"):
                c["ticker"] = (
                    ui.input(ROTULO_ATIVO.get(d.tipo, "Ativo"), value=d.ticker)
                    .props("outlined dense")
                    .classes("flex-1")
                    .mark(m("ativo-evento"))
                )
                if d.cv is not None and d.cv != SENTIDO.get(d.tipo, d.cv):
                    c["cv"] = (
                        ui.toggle({"C": "Soma", "V": "Tira"}, value=d.cv)
                        .props("dense")
                        .mark(m("cv-evento"))
                    )
            with ui.row().classes("w-full gap-2 flex-nowrap"):
                if d.tipo not in SEM_QUANTIDADE:
                    c["quantidade"] = (
                        ui.input(ROTULO_QUANTIDADE.get(d.tipo, "Quantidade"),
                                 value=_numero(d.quantidade))
                        .props("outlined dense input-class=text-right")
                        .classes("flex-1")
                        .mark(m("quantidade-evento"))
                    )  # fmt: skip
                sem_valor = (d.tipo in COM_DESTINO or d.tipo in (TipoEvento.TRANSFERENCIA,
                                                                 TipoEvento.DIREITO_NAO_EXERCIDO)
                             or (d.tipo in SEM_CUSTO and not d.valor))  # fmt: skip
                if not sem_valor:
                    c["valor"] = (
                        ui.input(ROTULO_VALOR.get(d.tipo, "Custo"), value=_numero(d.valor or None))
                        .props("outlined dense input-class=text-right")
                        .classes("flex-1")
                        .mark(m("valor-evento"))
                    )
                if d.tipo is TipoEvento.DIVIDENDO_EM_BENS:
                    c["valor_corrente"] = (
                        ui.input("Valor de mercado na entrega (total)",
                                 value=_numero(d.valor_corrente))
                        .props("outlined dense input-class=text-right")
                        .classes("flex-1")
                        .mark(m("valor-corrente-evento"))
                    )  # fmt: skip
            if d.tipo in COM_RECEBIDO:
                recebido, quanto = ROTULO_RECEBIDO.get(
                    d.tipo, ("Ativo recebido", "Quantidade recebida")
                )
                with ui.row().classes("w-full gap-2 flex-nowrap"):
                    c["ticker_destino"] = (
                        ui.input(recebido, value=d.ticker_destino or "")
                        .props("outlined dense")
                        .classes("flex-1")
                        .mark(m("recebido-evento"))
                    )
                    c["quantidade_destino"] = (
                        ui.input(quanto, value=_numero(d.quantidade_destino))
                        .props("outlined dense input-class=text-right")
                        .classes("flex-1")
                        .mark(m("qtd-recebida-evento"))
                    )
            if d.tipo is TipoEvento.INCORPORACAO:
                c["natureza"] = (
                    ui.select(
                        {n.value: r for n, r in NATUREZAS.items()},
                        value=(d.natureza or NaturezaEvento.INCORPORACAO_SOCIEDADE).value,
                        label="Natureza",
                        on_change=lambda e: aviso.set_visibility(
                            e.value == NaturezaEvento.INCORPORACAO_ACOES.value
                        ),
                    )
                    .props("outlined dense options-dense")
                    .classes("w-full")
                    .mark(m("natureza-evento"))
                )
                aviso = ui.label(AVISO_INCORPORACAO_ACOES).classes("coir-carimbo text-sm")
                aviso.set_visibility(d.natureza is NaturezaEvento.INCORPORACAO_ACOES)
            if d.tipo is TipoEvento.CISAO:
                self._campos_cisao(d)
            c["observacao"] = (
                ui.input("Observação (opcional)", value=d.observacao)
                .props("outlined dense")
                .classes("w-full")
                .mark(m("obs-evento"))
            )
            with ui.row().classes("w-full gap-2"):
                ui.button("Gravar", icon="save", on_click=self._gravar).props("unelevated").mark(
                    m("gravar-evento")
                )
                if form.evento_id is not None:
                    ui.button("Excluir", icon="delete", on_click=self._excluir).props(
                        "flat color=negative"
                    ).mark(m("excluir-evento"))
                if form.proposta is not None:
                    ignorada = form.proposta.situacao == "ignorada"
                    ui.button("Voltar a propor" if ignorada else "Ignorar proposta",
                              icon="undo" if ignorada else "block",
                              on_click=lambda: self._ignorar(not ignorada)).props("flat").mark(
                        m("ignorar-proposta"))  # fmt: skip
                if form.evento_id is not None and self.aba is not None:
                    ui.button("Novo", icon="add", on_click=self.novo).props("flat").mark(
                        m("novo-evento"))  # fmt: skip

    def _campos_cisao(self, d: DadosEvento) -> None:
        c, m, form, contas = self.form.campos, self.m, self.form, self.tela.contas
        c["percentual"] = (
            ui.input(
                "% do custo que vai para o ativo recebido",
                value=_numero(None if d.percentual_custo is None else d.percentual_custo * 100),
                placeholder="aviso aos acionistas, todas as casas",
            )
            .props("outlined dense input-class=text-right")
            .classes("w-full")
            .mark(m("percentual-evento"))
        )
        if form.evento_id is not None:
            return
        ui.label(
            "Posição em outras contas: lance todas juntas (o custo é repartido sobre a posição "
            "inteira)."
        ).classes("coir-subtitulo text-xs")
        for i, (conta, qtd, recebida) in enumerate(form.outras):
            with ui.row().classes("w-full gap-2 flex-nowrap"):
                c[f"conta_{i}"] = ui.select(contas, value=conta, label="Conta").props(
                    "outlined dense options-dense").classes("flex-1").mark(
                    m(f"conta-extra-{i}"))  # fmt: skip
                c[f"quantidade_{i}"] = ui.input("Qtd. de origem", value=qtd).props(
                    "outlined dense input-class=text-right").classes("w-[120px]").mark(
                    m(f"qtd-extra-{i}"))  # fmt: skip
                c[f"quantidade_destino_{i}"] = ui.input("Qtd. recebida", value=recebida).props(
                    "outlined dense input-class=text-right").classes("w-[120px]").mark(
                    m(f"qtd-recebida-extra-{i}"))  # fmt: skip
        ui.button("Outra conta", icon="add", on_click=self._mais_uma_conta).props(
            "flat dense").mark(m("outra-conta"))  # fmt: skip
