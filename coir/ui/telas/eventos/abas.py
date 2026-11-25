"""As abas da tela Eventos (quais tipos cada uma lança e lista) e as linhas das grades de
lançados."""

from collections.abc import Awaitable, Callable
from dataclasses import dataclass, field
from decimal import Decimal

from nicegui import ui

from coir.domain.fiscal import TipoEvento
from coir.reports.modelo import Coluna, Valor
from coir.servicos.estado import Estado
from coir.servicos.eventos import DadosEvento, EventoLancado, Proposta
from coir.ui.telas.eventos.textos import (
    COLUNAS_COMPRA,
    COM_DESTINO,
    NATUREZAS,
    NOS_LANCADOS,
    SUBSCRICOES,
    TIPOS,
    T,
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


def _linha_subscricao(e: EventoLancado) -> tuple[Valor, ...]:
    d = e.dados
    ativo = f"{d.ticker} → {d.ticker_destino}" if d.ticker_destino else d.ticker
    return (d.data, TIPOS[d.tipo], ativo, e.conta_nome, d.quantidade_destino or d.quantidade,
            d.valor or None)  # fmt: skip


def _linha_transferencia(e: EventoLancado) -> tuple[Valor, ...]:
    d = e.dados
    quanto = _numero(d.quantidade) or "toda a posição"
    return (d.data, d.ticker, e.conta_nome, e.conta_destino_nome or "", quanto)


def _linha_conversao(e: EventoLancado) -> tuple[Valor, ...]:
    d = e.dados
    natureza = (
        NATUREZAS[d.natureza]
        if d.natureza is not None and d.natureza in NATUREZAS
        else TIPOS[d.tipo]
    )
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
         TipoEvento.REDUCAO_CAPITAL, TipoEvento.LEILAO_FRACAO, TipoEvento.DIREITO_NAO_EXERCIDO,
         TipoEvento.DIVIDENDO_EM_BENS),
        "Outro evento",
        "Todos os eventos do titular. Clique num deles para corrigir ou excluir.",
        "Nenhum evento lançado.",
        (Coluna("Data", T.DATA), Coluna("Tipo"), Coluna("Ativo"), Coluna("Conta"),
         Coluna("Ações", T.QUANTIDADE), Coluna("Valor", T.DINHEIRO), Coluna("Recebe"),
         Coluna("Origem")),
        _linha_lancado,
    ),
    Aba(
        "subscricoes", "Subscrições", "Subscrições", SUBSCRICOES,
        (TipoEvento.EXERCICIO_DIREITO, TipoEvento.DIREITO_RECEBIDO, TipoEvento.SUBSCRICAO),
        "Nova subscrição",
        "Direitos de subscrição que você recebeu e ações novas que comprou com eles. O custo do "
        "direito comprado em bolsa entra no custo das ações subscritas.",
        "Nenhuma subscrição lançada.",
        (Coluna("Data", T.DATA), Coluna("O que houve"), Coluna("Ativo"), Coluna("Conta"),
         Coluna("Quantidade", T.QUANTIDADE), Coluna("Valor pago", T.DINHEIRO)),
        _linha_subscricao,
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
