"""Tela 7 — Eventos corporativos: no topo, os eventos que a movimentação da B3 sugere (lançar ou
ignorar); abaixo, os eventos do titular e o painel para incluir ou editar. Cada evento gravado é
conferido pelo motor (sem posição, quantidade maior que a posição... não é gravado)."""

from collections.abc import Callable
from dataclasses import dataclass, field, replace
from datetime import date
from decimal import Decimal

from nicegui import run, ui

from coir.domain.fiscal import NaturezaEvento, TipoEvento
from coir.erros import CoirError
from coir.reports.modelo import Coluna, TipoColuna, Valor
from coir.servicos import eventos
from coir.servicos.estado import Estado
from coir.servicos.eventos import AVISO_INCORPORACAO_ACOES, DadosEvento, EventoLancado, Proposta
from coir.ui import tabela
from coir.ui.entrada import valor_br

T = TipoColuna
TIPOS = {
    TipoEvento.DESDOBRAMENTO: "Desdobramento",
    TipoEvento.GRUPAMENTO: "Grupamento",
    TipoEvento.BONIFICACAO: "Bonificação",
    TipoEvento.SUBSCRICAO: "Subscrição",
    TipoEvento.CUSTODIA_INICIAL: "Custódia inicial (31/12)",
    TipoEvento.REDUCAO_CAPITAL: "Redução de capital",
    TipoEvento.TRANSFERENCIA: "Transferência de custódia",
    TipoEvento.INCORPORACAO: "Incorporação / conversão / ajuste de série",
    TipoEvento.CISAO: "Cisão",
    TipoEvento.LEILAO_FRACAO: "Leilão de fração",
}
NATUREZAS = {
    NaturezaEvento.INCORPORACAO_SOCIEDADE: "Incorporação de sociedade",
    NaturezaEvento.INCORPORACAO_ACOES: "Incorporação de ações (art. 252)",
    NaturezaEvento.FUSAO: "Fusão",
    NaturezaEvento.CONVERSAO: "Conversão de classe (ex.: PN em ON)",
    NaturezaEvento.AJUSTE_SERIE: "Ajuste de série de opção pela B3",
}
AJUDA = {
    TipoEvento.DESDOBRAMENTO: "C com as ações recebidas, custo zero.",
    TipoEvento.GRUPAMENTO: "V com as ações que deixam de existir, custo zero.",
    TipoEvento.BONIFICACAO: "C com as ações recebidas e o custo atribuído pela companhia.",
    TipoEvento.SUBSCRICAO: "C com as ações subscritas e o valor pago.",
    TipoEvento.CUSTODIA_INICIAL: "Para quem não tem as notas antigas: C com a quantidade e o custo "
                                 "em 31/12 do Bens e Direitos da última declaração (um por ativo "
                                 "e conta). Opções e futuros em aberto pedem as notas.",
    TipoEvento.REDUCAO_CAPITAL: "V com o valor restituído: reduz o custo, sem resultado.",
    TipoEvento.TRANSFERENCIA: "Move a posição para outra conta do titular com o mesmo PM "
                              "(quantidade vazia = toda).",
    TipoEvento.INCORPORACAO: "A quantidade de origem sai sem resultado e o ativo recebido entra "
                             "com o mesmo custo (IN 1.585 art. 58 §6º). Série de opção: herda o "
                             "vencimento.",
    TipoEvento.CISAO: "O percentual do custo divulgado pela companhia vai para o ativo recebido; "
                      "a quantidade de origem é a posição da conta (lance cada conta).",
    TipoEvento.LEILAO_FRACAO: "Venda da fração pelo valor líquido creditado (data do crédito); "
                              "entra no limite de isenção do mês.",
}  # fmt: skip
COM_CV = frozenset(TIPOS) - eventos.SEM_CV
COM_DESTINO = frozenset({TipoEvento.INCORPORACAO, TipoEvento.CISAO})
COLUNAS = (
    Coluna("Data", T.DATA),
    Coluna("Tipo"),
    Coluna("Ativo"),
    Coluna("Conta"),
    Coluna("C/V"),
    Coluna("Qtd.", T.QUANTIDADE),
    Coluna("Valor", T.DINHEIRO),
    Coluna("Recebe"),
    Coluna("Origem"),
)
COLUNAS_PROPOSTAS = (
    Coluna("Data na B3", T.DATA),
    Coluna("Movimentação"),
    Coluna("Ativo"),
    Coluna("Corretora"),
    Coluna("Qtd.", T.QUANTIDADE),
    Coluna("Valor", T.DINHEIRO),
    Coluna("Situação"),
)


def _numero(v: Decimal | None) -> str:
    return "" if v is None else format(v.normalize(), "f").replace(".", ",")


def _recebe(d: DadosEvento, contas: dict[int, str]) -> str:
    if d.tipo is TipoEvento.TRANSFERENCIA:
        return f"→ {contas.get(d.conta_destino or 0, '')}"
    if d.ticker_destino:
        extra = "" if d.percentual_custo is None else f" ({_numero(d.percentual_custo * 100)}%)"
        return f"{_numero(d.quantidade_destino)} {d.ticker_destino}{extra}"
    return ""


def valores_lancados(lista: list[EventoLancado], contas: dict[int, str]) -> list[tuple[Valor, ...]]:
    return [
        (e.dados.data, TIPOS[e.dados.tipo], e.dados.ticker, e.conta_nome, e.dados.cv or "",
         e.dados.quantidade, e.dados.valor or None, _recebe(e.dados, contas), e.dados.origem)
        for e in lista
    ]  # fmt: skip


@dataclass
class _Form:
    """O que está no painel (novo evento, edição de um lançado ou proposta da B3)."""

    dados: DadosEvento
    evento_id: int | None = None
    proposta: Proposta | None = None
    campos: dict[str, ui.element] = field(default_factory=dict[str, ui.element])
    # cisão: as outras contas da mesma posição (conta, quantidade de origem, quantidade recebida)
    outras: list[tuple[int | None, str, str]] = field(
        default_factory=list[tuple[int | None, str, str]]
    )


def _vazio(contas: dict[int, str], dia: date) -> DadosEvento:
    return DadosEvento(tipo=TipoEvento.DESDOBRAMENTO, data=dia,
                       conta=next(iter(contas), None), ticker="", cv="C")  # fmt: skip


def montar(estado: Estado, ao_mudar: Callable[[], object]) -> None:
    ui.label("Eventos corporativos").classes("coir-titulo-pagina")
    ui.label(
        "Desdobramentos, grupamentos, bonificações, subscrições, reduções de capital, "
        "transferências, incorporações, cisões e leilões de fração. A data é a data ex (primeiro "
        "pregão sem o direito); a movimentação da B3 traz a data do crédito. Cada evento é "
        "conferido com as operações antes de ser gravado."
    ).classes("coir-subtitulo")
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
    form = _Form(_vazio(contas, estado.fim_do_ano()))
    todas = {"ver": False}

    secao_propostas = ui.column().classes("w-full gap-2")
    with ui.row().classes("w-full items-start gap-4 flex-nowrap"):
        lista = ui.column().classes("min-w-[560px] flex-[1.4] gap-2")
        painel = ui.column().classes("coir-folha min-w-[380px] flex-1 gap-2").mark("painel-evento")

    async def recarregar() -> None:
        try:
            lancados = await run.io_bound(eventos.listar, estado.banco, titular)
            # a custódia inicial fica em Cadastros > Saldos iniciais
            lancados = [
                e for e in lancados or [] if e.dados.tipo is not TipoEvento.CUSTODIA_INICIAL
            ]
            propostas = await run.io_bound(estado.propostas)
        except CoirError as erro:
            lista.clear()
            with lista:
                ui.label(f"Não foi possível ler os eventos: {erro}").classes("coir-negativo")
            return
        _propostas(propostas or [])
        _lista(lancados or [])

    def _propostas(propostas: list[Proposta]) -> None:
        secao_propostas.clear()
        pendentes = [p for p in propostas if p.situacao == "pendente"]
        mostradas = propostas if todas["ver"] else pendentes
        with secao_propostas:
            with ui.row().classes("w-full items-center gap-3"):
                ui.label(f"Propostas da B3 ({len(pendentes)})").classes("coir-titulo text-lg")
                ui.checkbox("mostrar lançadas e ignoradas", value=todas["ver"],
                            on_change=lambda e: ver_todas(bool(e.value), propostas)
                            ).mark("ver-propostas")  # fmt: skip
            if not propostas:
                ui.label("Importe a movimentação da B3 (Importar > Planilhas da B3) para ver os "
                         "eventos sugeridos.").classes("coir-vazio w-full")  # fmt: skip
            elif not mostradas:
                ui.label("Nenhuma proposta pendente: os eventos da movimentação já estão "
                         "lançados ou ignorados.").classes("coir-vazio w-full")  # fmt: skip
            else:
                tabela.grade(
                    COLUNAS_PROPOSTAS,
                    [(p.data, p.movimentacao, p.codigo, p.corretora or "", p.quantidade, p.valor,
                      p.situacao) for p in mostradas],
                    chaves=list(range(len(mostradas))), altura="h-[220px]",
                    ao_clicar=lambda i: abrir_proposta(mostradas[int(str(i))]),
                ).mark("grade-propostas")  # fmt: skip
                ui.label("Clique numa proposta para conferir e lançar (ou ignorar).").classes(
                    "coir-subtitulo text-xs"
                )

    def ver_todas(sim: bool, propostas: list[Proposta]) -> None:
        todas["ver"] = sim
        _propostas(propostas)

    def _lista(lancados: list[EventoLancado]) -> None:
        lista.clear()
        with lista:
            ui.label(f"Eventos lançados ({len(lancados)})").classes("coir-titulo text-lg")
            if not lancados:
                ui.label("Nenhum evento lançado.").classes("coir-vazio w-full")
                return
            tabela.grade(
                COLUNAS, valores_lancados(lancados, contas),
                chaves=list(range(len(lancados))), altura="h-[520px]",
                ao_clicar=lambda i: abrir_evento(lancados[int(str(i))]),
            ).mark("grade-eventos")  # fmt: skip

    def abrir_evento(e: EventoLancado) -> None:
        form.dados, form.evento_id, form.proposta, form.outras = e.dados, e.id, None, []
        desenhar()

    def abrir_proposta(p: Proposta) -> None:
        form.dados, form.evento_id, form.proposta, form.outras = p.dados, None, p, []
        desenhar()

    def novo() -> None:
        form.dados, form.evento_id, form.proposta = _vazio(contas, estado.fim_do_ano()), None, None
        form.outras = []
        desenhar()

    def ler() -> DadosEvento | None:
        """Os campos do painel -> DadosEvento (None se algum número está ilegível)."""
        c = form.campos

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

        d = form.dados
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
                cv=(texto("cv") or None) if d.tipo in COM_CV else None,
                quantidade=numero("quantidade"),
                valor=numero("valor") or Decimal(0),
                conta_destino=getattr(c["destino"], "value", None) if "destino" in c else None,
                ticker_destino=texto("ticker_destino") or None,
                quantidade_destino=numero("quantidade_destino"),
                natureza=(NaturezaEvento(texto("natureza")) if texto("natureza") else None),
                percentual_custo=None if percentual is None else percentual / 100,
                observacao=texto("observacao"),
            )
        except ValueError:
            return None

    def ler_outras() -> list[tuple[int | None, str, str]]:
        c = form.campos
        return [
            (getattr(c[f"conta_{i}"], "value", None),
             str(getattr(c[f"quantidade_{i}"], "value", "") or "").strip(),
             str(getattr(c[f"quantidade_destino_{i}"], "value", "") or "").strip())
            for i in range(len(form.outras))
        ]  # fmt: skip

    def mais_uma_conta() -> None:
        form.dados = ler() or form.dados
        form.outras = [*ler_outras(), (None, "", "")]
        desenhar()

    def lote(d: DadosEvento) -> list[tuple[DadosEvento, int | None]] | None:
        """O evento do painel e, na cisão, o das outras contas (mesmos dados, outra conta)."""
        saida: list[tuple[DadosEvento, int | None]] = [(d, form.evento_id)]
        for conta, qtd, recebida in ler_outras() if d.tipo is TipoEvento.CISAO else []:
            q, r = valor_br(qtd) if qtd else None, valor_br(recebida) if recebida else None
            if conta is None or q is None or r is None:
                ui.notify("Outra conta: informe a conta e as duas quantidades.", type="warning")
                return None
            saida.append((replace(d, conta=conta, quantidade=q, quantidade_destino=r), None))
        return saida

    def mudar_tipo(valor: object) -> None:
        atual = ler() or form.dados
        tipo = TipoEvento(str(valor))
        form.dados = replace(
            atual,
            tipo=tipo,
            cv=("C" if tipo in COM_CV and not atual.cv else atual.cv if tipo in COM_CV else None),
        )
        desenhar()

    async def salvar() -> None:
        d = ler()
        pedidos = None if d is None else lote(d)
        if pedidos is None:
            return
        try:
            ids = await run.io_bound(eventos.salvar_lote, estado.banco, titular, pedidos,
                                     estado.refs, estado.fiscais)  # fmt: skip
            novo_id = ", ".join(str(i) for i in ids or [])
        except CoirError as erro:
            ui.notify(str(erro), type="negative", multi_line=True, timeout=12000)
            return
        estado.invalidar()
        ao_mudar()
        ui.notify(f"Evento gravado (id {novo_id}). Carteira e apuração refeitas.",
                  type="positive")  # fmt: skip
        novo()
        await recarregar()

    async def excluir() -> None:
        if form.evento_id is None:
            return
        try:
            await run.io_bound(
                eventos.excluir, estado.banco, titular, form.evento_id, estado.refs, estado.fiscais
            )
        except CoirError as erro:
            ui.notify(str(erro), type="negative")
            return
        estado.invalidar()
        ao_mudar()
        ui.notify("Evento excluído.", type="positive")
        novo()
        await recarregar()

    async def ignorar(sim: bool) -> None:
        if form.proposta is None:
            return
        await run.io_bound(eventos.ignorar, estado.banco, titular, form.proposta.chave,
                           estado.agora(), sim)  # fmt: skip
        novo()
        await recarregar()

    def desenhar() -> None:
        d = form.dados
        painel.clear()
        form.campos = {}
        c = form.campos
        with painel:
            if form.proposta is not None:
                titulo = f"Proposta da B3: {form.proposta.movimentacao} {form.proposta.codigo}"
            elif form.evento_id is not None:
                titulo = f"Editar evento {form.evento_id}"
            else:
                titulo = "Novo evento"
            ui.label(titulo).classes("coir-titulo text-lg").mark("titulo-painel")
            if form.proposta is not None:
                ui.label(form.proposta.nota).classes("coir-carimbo text-sm")
            ui.select(
                {
                    t.value: r
                    for t, r in TIPOS.items()
                    if t is not TipoEvento.CUSTODIA_INICIAL or d.tipo is t
                },
                value=d.tipo.value,
                label="Tipo",
                on_change=lambda e: mudar_tipo(e.value),
            ).props("outlined dense options-dense").classes("w-full").mark("tipo-evento")
            ui.label(AJUDA[d.tipo]).classes("coir-subtitulo text-xs")
            with ui.row().classes("w-full gap-2 flex-nowrap"):
                c["data"] = (
                    ui.input("Data (ex)", value=d.data.isoformat())
                    .props("outlined dense type=date")
                    .classes("flex-1")
                    .mark("data-evento")
                )
                c["conta"] = (
                    ui.select(contas, value=d.conta, label="Conta")
                    .props("outlined dense options-dense")
                    .classes("flex-1")
                    .mark("conta-evento")
                )
            with ui.row().classes("w-full gap-2 flex-nowrap"):
                c["ticker"] = (
                    ui.input("Ativo", value=d.ticker)
                    .props("outlined dense")
                    .classes("flex-1")
                    .mark("ativo-evento")
                )
                if d.tipo in COM_CV:
                    c["cv"] = (
                        ui.toggle(["C", "V"], value=d.cv or "C").props("dense").mark("cv-evento")
                    )
            rotulo_qtd = {
                TipoEvento.INCORPORACAO: "Quantidade de origem",
                TipoEvento.CISAO: "Quantidade de origem (posição da conta)",
                TipoEvento.LEILAO_FRACAO: "Fração vendida",
            }.get(d.tipo, "Quantidade")
            with ui.row().classes("w-full gap-2 flex-nowrap"):
                c["quantidade"] = (
                    ui.input(rotulo_qtd, value=_numero(d.quantidade))
                    .props("outlined dense input-class=text-right")
                    .classes("flex-1")
                    .mark("quantidade-evento")
                )
                if d.tipo not in COM_DESTINO and d.tipo is not TipoEvento.TRANSFERENCIA:
                    rotulo_valor = {TipoEvento.BONIFICACAO: "Custo atribuído",
                                    TipoEvento.SUBSCRICAO: "Valor pago",
                                    TipoEvento.CUSTODIA_INICIAL: "Custo de aquisição",
                                    TipoEvento.REDUCAO_CAPITAL: "Valor restituído",
                                    TipoEvento.LEILAO_FRACAO: "Valor líquido recebido",
                                    }.get(d.tipo, "Custo")  # fmt: skip
                    c["valor"] = (
                        ui.input(rotulo_valor, value=_numero(d.valor or None))
                        .props("outlined dense input-class=text-right")
                        .classes("flex-1")
                        .mark("valor-evento")
                    )
            if d.tipo is TipoEvento.TRANSFERENCIA:
                c["destino"] = (
                    ui.select(
                        {k: v for k, v in contas.items() if k != d.conta},
                        value=d.conta_destino,
                        label="Conta de destino",
                    )
                    .props("outlined dense options-dense")
                    .classes("w-full")
                    .mark("destino-evento")
                )
            if d.tipo in COM_DESTINO:
                with ui.row().classes("w-full gap-2 flex-nowrap"):
                    c["ticker_destino"] = (
                        ui.input("Ativo recebido", value=d.ticker_destino or "")
                        .props("outlined dense")
                        .classes("flex-1")
                        .mark("recebido-evento")
                    )
                    c["quantidade_destino"] = (
                        ui.input("Quantidade recebida", value=_numero(d.quantidade_destino))
                        .props("outlined dense input-class=text-right")
                        .classes("flex-1")
                        .mark("qtd-recebida-evento")
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
                    .mark("natureza-evento")
                )
                aviso = ui.label(AVISO_INCORPORACAO_ACOES).classes("coir-carimbo text-sm")
                aviso.set_visibility(d.natureza is NaturezaEvento.INCORPORACAO_ACOES)
            if d.tipo is TipoEvento.CISAO:
                c["percentual"] = (
                    ui.input(
                        "% do custo que vai para o ativo recebido",
                        value=_numero(
                            None if d.percentual_custo is None else d.percentual_custo * 100
                        ),
                        placeholder="aviso aos acionistas, todas as casas",
                    )
                    .props("outlined dense input-class=text-right")
                    .classes("w-full")
                    .mark("percentual-evento")
                )
                if form.evento_id is None:
                    ui.label(
                        "Posição em outras contas: lance todas juntas (o custo é repartido sobre "
                        "a posição inteira)."
                    ).classes("coir-subtitulo text-xs")
                    for i, (conta, qtd, recebida) in enumerate(form.outras):
                        with ui.row().classes("w-full gap-2 flex-nowrap"):
                            c[f"conta_{i}"] = ui.select(contas, value=conta, label="Conta").props(
                                "outlined dense options-dense").classes("flex-1").mark(
                                f"conta-extra-{i}")  # fmt: skip
                            c[f"quantidade_{i}"] = ui.input("Qtd. de origem", value=qtd).props(
                                "outlined dense input-class=text-right").classes("w-[120px]"
                                ).mark(f"qtd-extra-{i}")  # fmt: skip
                            c[f"quantidade_destino_{i}"] = ui.input(
                                "Qtd. recebida", value=recebida).props(
                                "outlined dense input-class=text-right").classes("w-[120px]"
                                ).mark(f"qtd-recebida-extra-{i}")  # fmt: skip
                    ui.button("Outra conta", icon="add", on_click=mais_uma_conta).props(
                        "flat dense").mark("outra-conta")  # fmt: skip
            c["observacao"] = (
                ui.input("Observação", value=d.observacao)
                .props("outlined dense")
                .classes("w-full")
                .mark("obs-evento")
            )
            with ui.row().classes("w-full gap-2"):
                ui.button("Gravar", icon="save", on_click=salvar).props("unelevated").mark(
                    "gravar-evento"
                )
                if form.evento_id is not None:
                    ui.button("Excluir", icon="delete", on_click=excluir).props(
                        "flat color=negative"
                    ).mark("excluir-evento")
                if form.proposta is not None:
                    ignorada = form.proposta.situacao == "ignorada"
                    ui.button("Voltar a propor" if ignorada else "Ignorar proposta",
                              icon="undo" if ignorada else "block",
                              on_click=lambda: ignorar(not ignorada)).props("flat").mark(
                        "ignorar-proposta")  # fmt: skip
                if form.evento_id is not None or form.proposta is not None:
                    ui.button("Novo", icon="add", on_click=novo).props("flat").mark("novo-evento")

    desenhar()
    ui.timer(0.01, recarregar, once=True)
