"""Tela 7 — Eventos: o que muda a posição sem passar por uma nota. Na primeira aba, o que a B3
sugere (conferir, lançar ou ignorar); nas outras, os lançados e um formulário curto por tarefa. Cada
evento gravado é conferido pelo motor (sem posição, quantidade maior que a posição... não é
gravado)."""

from collections.abc import Awaitable, Callable
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
    TipoEvento.OFERTA_PUBLICA: "Oferta pública (IPO)",
    TipoEvento.CUSTODIA_INICIAL: "Custódia inicial (31/12)",
    TipoEvento.REDUCAO_CAPITAL: "Redução de capital",
    TipoEvento.TRANSFERENCIA: "Transferência de custódia",
    TipoEvento.INCORPORACAO: "Incorporação / conversão / ajuste de série",
    TipoEvento.CISAO: "Cisão",
    TipoEvento.LEILAO_FRACAO: "Leilão de fração",
    TipoEvento.DIREITO_NAO_EXERCIDO: "Direito não exercido",
}
NATUREZAS = {
    NaturezaEvento.CONVERSAO: "Conversão de classe (ex.: PN em ON)",
    NaturezaEvento.INCORPORACAO_SOCIEDADE: "Incorporação de sociedade",
    NaturezaEvento.INCORPORACAO_ACOES: "Incorporação de ações (art. 252)",
    NaturezaEvento.FUSAO: "Fusão",
    NaturezaEvento.AJUSTE_SERIE: "Ajuste de série de opção pela B3",
}
AJUDA = {
    TipoEvento.DESDOBRAMENTO: "Informe as ações recebidas. O custo total não muda e o preço médio "
                              "cai.",
    TipoEvento.GRUPAMENTO: "Informe as ações que deixaram de existir. O custo total não muda e o "
                           "preço médio sobe.",
    TipoEvento.BONIFICACAO: "Informe as ações recebidas e o custo atribuído pela companhia, que "
                            "está no aviso aos acionistas.",
    TipoEvento.SUBSCRICAO: "Informe as ações subscritas e o valor pago.",
    TipoEvento.OFERTA_PUBLICA: "Informe as ações compradas na oferta e o valor pago (quantidade "
                               "vezes preço). Entra no preço médio como qualquer compra.",
    TipoEvento.CUSTODIA_INICIAL: "Para quem não tem as notas antigas: a quantidade e o custo em "
                                 "31/12 do Bens e Direitos da última declaração (um por ativo e "
                                 "conta). Opções e futuros em aberto pedem as notas.",
    TipoEvento.REDUCAO_CAPITAL: "Informe o valor restituído pela companhia. Ele reduz o custo, sem "
                                "resultado.",
    TipoEvento.TRANSFERENCIA: "Move a posição para outra conta do titular, com o mesmo preço "
                              "médio. Deixe a quantidade vazia para mover toda a posição.",
    TipoEvento.INCORPORACAO: "As ações de origem saem sem resultado e as recebidas entram com o "
                             "mesmo custo (IN 1.585 art. 58 §6º). Série de opção: herda o "
                             "vencimento.",
    TipoEvento.CISAO: "O percentual do custo divulgado pela companhia vai para o ativo recebido. "
                      "A quantidade de origem é a posição da conta (lance cada conta).",
    TipoEvento.LEILAO_FRACAO: "Venda da fração pelo valor líquido creditado, na data do crédito. "
                              "Entra no limite de isenção do mês.",
    TipoEvento.DIREITO_NAO_EXERCIDO: "Direito, recibo ou bônus de subscrição comprado que venceu "
                                     "sem exercício nem venda: o custo pago vira perda em "
                                     "operações comuns, fora do limite de isenção. Use o fim do "
                                     "prazo de exercício. Sem quantidade, baixa todo o saldo; "
                                     "os recebidos sem custo não dão resultado.",
}  # fmt: skip
ROTULO_DATA = {
    TipoEvento.SUBSCRICAO: "Data da subscrição",
    TipoEvento.OFERTA_PUBLICA: "Data da liquidação",
    TipoEvento.TRANSFERENCIA: "Data da transferência",
    TipoEvento.LEILAO_FRACAO: "Data do crédito",
    TipoEvento.DIREITO_NAO_EXERCIDO: "Fim do prazo de exercício",
    TipoEvento.CUSTODIA_INICIAL: "Data",
}
ROTULO_QUANTIDADE = {
    TipoEvento.DESDOBRAMENTO: "Ações recebidas",
    TipoEvento.GRUPAMENTO: "Ações que deixaram de existir",
    TipoEvento.BONIFICACAO: "Ações recebidas",
    TipoEvento.SUBSCRICAO: "Ações subscritas",
    TipoEvento.OFERTA_PUBLICA: "Ações compradas",
    TipoEvento.INCORPORACAO: "Quantidade de origem",
    TipoEvento.CISAO: "Quantidade de origem (posição da conta)",
    TipoEvento.LEILAO_FRACAO: "Fração vendida",
    TipoEvento.DIREITO_NAO_EXERCIDO: "Direitos que venceram (vazio: todo o saldo)",
}
ROTULO_VALOR = {
    TipoEvento.BONIFICACAO: "Custo atribuído",
    TipoEvento.SUBSCRICAO: "Valor pago",
    TipoEvento.OFERTA_PUBLICA: "Valor pago",
    TipoEvento.CUSTODIA_INICIAL: "Custo de aquisição",
    TipoEvento.REDUCAO_CAPITAL: "Valor restituído",
    TipoEvento.LEILAO_FRACAO: "Valor líquido recebido",
}
# se o evento soma ou tira da posição: só aparece para escolher quando foge do normal (ajuste de
# série de opção vendida, evento antigo lançado ao contrário)
SENTIDO = {
    TipoEvento.DESDOBRAMENTO: "C",
    TipoEvento.GRUPAMENTO: "V",
    TipoEvento.BONIFICACAO: "C",
    TipoEvento.SUBSCRICAO: "C",
    TipoEvento.OFERTA_PUBLICA: "C",
    TipoEvento.CUSTODIA_INICIAL: "C",
    TipoEvento.REDUCAO_CAPITAL: "V",
}
SEM_CUSTO = frozenset({TipoEvento.DESDOBRAMENTO, TipoEvento.GRUPAMENTO})
COM_OUTRA_CLASSE = frozenset({TipoEvento.BONIFICACAO, TipoEvento.SUBSCRICAO})
COM_DESTINO = frozenset({TipoEvento.INCORPORACAO, TipoEvento.CISAO})
NOS_LANCADOS = frozenset(TIPOS) - {TipoEvento.CUSTODIA_INICIAL}  # a custódia fica em Cadastros

COLUNAS_PROPOSTAS = (
    Coluna("Data na B3", T.DATA),
    Coluna("Evento"),
    Coluna("Ativo"),
    Coluna("Corretora"),
    Coluna("Qtd.", T.QUANTIDADE),
    Coluna("Valor", T.DINHEIRO),
    Coluna("Situação"),
)
COLUNAS_SEGURAS = (
    Coluna("Data ex", T.DATA),
    Coluna("Evento"),
    Coluna("Ativo"),
    Coluna("Corretora"),
    Coluna("Ações", T.QUANTIDADE),
    Coluna("Custo total", T.DINHEIRO),
    Coluna("R$ por ação", T.QUANTIDADE),
    Coluna("Fonte do custo"),
)
COLUNAS_COMPRA = (
    Coluna("Data", T.DATA),
    Coluna("Ativo"),
    Coluna("Conta"),
    Coluna("Ações", T.QUANTIDADE),
    Coluna("Valor pago", T.DINHEIRO),
)


def _numero(v: Decimal | None) -> str:
    return "" if v is None else format(v.normalize(), "f").replace(".", ",")


def _quantidade(d: DadosEvento) -> Decimal | None:
    """Com sinal: o que sai da posição aparece negativo."""
    return None if d.quantidade is None else (-d.quantidade if d.cv == "V" else d.quantidade)


def _recebe(e: EventoLancado) -> str:
    d = e.dados
    if d.tipo is TipoEvento.TRANSFERENCIA:
        return f"→ {e.conta_destino_nome or ''}"
    if d.ticker_destino:
        extra = "" if d.percentual_custo is None else f" ({_numero(d.percentual_custo * 100)}%)"
        return f"{_numero(d.quantidade_destino)} {d.ticker_destino}{extra}"
    return ""


def _linha_lancado(e: EventoLancado) -> tuple[Valor, ...]:
    d = e.dados
    return (d.data, TIPOS[d.tipo], d.ticker, e.conta_nome, _quantidade(d), d.valor or None,
            _recebe(e), d.origem)  # fmt: skip


def _linha_compra(e: EventoLancado) -> tuple[Valor, ...]:
    d = e.dados
    return (d.data, d.ticker, e.conta_nome, d.quantidade, d.valor or None)


def _linha_transferencia(e: EventoLancado) -> tuple[Valor, ...]:
    d = e.dados
    quanto = _numero(d.quantidade) or "toda a posição"
    return (d.data, d.ticker, e.conta_nome, e.conta_destino_nome or "", quanto)


def _linha_conversao(e: EventoLancado) -> tuple[Valor, ...]:
    d = e.dados
    natureza = TIPOS[d.tipo] if d.natureza is None else NATUREZAS[d.natureza]
    return (d.data, f"{d.ticker} → {d.ticker_destino or ''}", e.conta_nome, d.quantidade,
            d.quantidade_destino, natureza)  # fmt: skip


@dataclass(frozen=True)
class Aba:
    """Uma tarefa: a lista dos eventos dela e o formulário só com os campos que ela pede."""

    chave: str
    rotulo: str
    titulo: str
    tipos: frozenset[TipoEvento]
    novos: tuple[TipoEvento, ...]
    titulo_novo: str
    explicacao: str
    vazio: str
    colunas: tuple[Coluna, ...]
    linha: Callable[[EventoLancado], tuple[Valor, ...]]


ABAS = (
    Aba(
        "lancados", "Lançados", "Eventos lançados", NOS_LANCADOS,
        (TipoEvento.DESDOBRAMENTO, TipoEvento.GRUPAMENTO, TipoEvento.BONIFICACAO,
         TipoEvento.REDUCAO_CAPITAL, TipoEvento.LEILAO_FRACAO, TipoEvento.DIREITO_NAO_EXERCIDO),
        "Outro evento",
        "Todos os eventos do titular. Clique num deles para corrigir ou excluir.",
        "Nenhum evento lançado.",
        (Coluna("Data", T.DATA), Coluna("Tipo"), Coluna("Ativo"), Coluna("Conta"),
         Coluna("Ações", T.QUANTIDADE), Coluna("Valor", T.DINHEIRO), Coluna("Recebe"),
         Coluna("Origem")),
        _linha_lancado,
    ),
    Aba(
        "subscricoes", "Subscrições", "Subscrições", frozenset({TipoEvento.SUBSCRICAO}),
        (TipoEvento.SUBSCRICAO,), "Nova subscrição",
        "Ações novas que você decidiu comprar pelo direito de acionista.",
        "Nenhuma subscrição lançada.", COLUNAS_COMPRA, _linha_compra,
    ),
    Aba(
        "transferencias", "Transferências", "Transferências",
        frozenset({TipoEvento.TRANSFERENCIA}), (TipoEvento.TRANSFERENCIA,), "Nova transferência",
        "Ativos que você levou de uma corretora para outra. Não muda o preço médio nem gera "
        "imposto.",
        "Nenhuma transferência lançada.",
        (Coluna("Data", T.DATA), Coluna("Ativo"), Coluna("Origem"), Coluna("Destino"),
         Coluna("Quantidade")),
        _linha_transferencia,
    ),
    Aba(
        "conversoes", "Conversões", "Conversões e incorporações",
        frozenset(COM_DESTINO), (TipoEvento.INCORPORACAO, TipoEvento.CISAO), "Nova conversão",
        "Quando um ativo vira outro: conversão de classe, incorporação, fusão ou cisão.",
        "Nenhuma conversão lançada.",
        (Coluna("Data", T.DATA), Coluna("Troca"), Coluna("Conta"), Coluna("Sai", T.QUANTIDADE),
         Coluna("Entra", T.QUANTIDADE), Coluna("O que houve")),
        _linha_conversao,
    ),
    Aba(
        "ofertas", "Ofertas (IPO)", "Ofertas públicas", frozenset({TipoEvento.OFERTA_PUBLICA}),
        (TipoEvento.OFERTA_PUBLICA,), "Nova oferta",
        "Compras em IPO ou follow-on. Elas não aparecem em nota de corretagem.",
        "Nenhuma oferta lançada.", COLUNAS_COMPRA, _linha_compra,
    ),
)  # fmt: skip


@dataclass
class _Form:
    """O que está no painel (novo evento, edição de um lançado ou proposta da B3)."""

    dados: DadosEvento | None
    evento_id: int | None = None
    proposta: Proposta | None = None
    campos: dict[str, ui.element] = field(default_factory=dict[str, ui.element])
    # cisão: as outras contas da mesma posição (conta, quantidade de origem, quantidade recebida)
    outras: list[tuple[int | None, str, str]] = field(
        default_factory=list[tuple[int | None, str, str]]
    )


@dataclass
class _Tela:
    estado: Estado
    titular: int
    contas: dict[int, str]
    ao_mudar: Callable[[], object]
    recarregar: Callable[[], Awaitable[None]]


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
        self.novo()
        await t.recarregar()

    async def _ignorar(self, sim: bool) -> None:
        t = self.tela
        if self.form.proposta is None:
            return
        await run.io_bound(eventos.ignorar, t.estado.banco, t.titular, self.form.proposta.chave,
                           t.estado.agora(), sim)  # fmt: skip
        t.ao_mudar()  # o sino do topo conta as propostas pendentes
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
                    ui.input("Ativo", value=d.ticker)
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
            if d.tipo in COM_DESTINO:
                with ui.row().classes("w-full gap-2 flex-nowrap"):
                    c["ticker_destino"] = (
                        ui.input("Ativo recebido", value=d.ticker_destino or "")
                        .props("outlined dense")
                        .classes("flex-1")
                        .mark(m("recebido-evento"))
                    )
                    c["quantidade_destino"] = (
                        ui.input("Quantidade recebida", value=_numero(d.quantidade_destino))
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
        saidas = t.estado.ativos_que_sairam()
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
        self.painel.novo()
        estado.invalidar()
        t.ao_mudar()
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
        self.painel.novo()
        t.estado.invalidar()
        t.ao_mudar()
        ui.notify("Lançamento desfeito.", type="positive")
        await t.recarregar()


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
            if conferir.lista.is_deleted:
                return  # o usuário saiu da tela durante a leitura
        except CoirError as erro:
            conferir.lista.clear()
            with conferir.lista:
                ui.label(f"Não foi possível ler os eventos: {erro}").classes("coir-negativo")
            return
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
