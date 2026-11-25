"""O evento como o usuário lança (DadosEvento), a validação, a chave que o identifica e as
consultas simples (lançados, ignorados, contas)."""

import hashlib
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

from sqlalchemy import select
from sqlalchemy.orm import Session, selectinload

from coir.db import repositorio as repo
from coir.db.modelos import (
    Conta,
    EventoCorporativo,
    PropostaIgnorada,
    TickerHistorico,
)
from coir.db.sessao import conectar
from coir.domain.fiscal import (
    TIPOS_COM_DESTINO,
    NaturezaEvento,
    TipoEvento,
)
from coir.domain.referencias import Referencias
from coir.erros import ImportacaoError
from coir.servicos.apuracao import Tickers

COLUNAS = ("data", "corretora", "conta", "ticker", "tipo", "cv", "quantidade", "valor",
           "corretora_destino", "conta_destino", "observacao", "origem")  # fmt: skip


EXTRAS = ("ticker_destino", "quantidade_destino", "natureza", "percentual_custo",
          "valor_corrente")  # opcionais  # fmt: skip


NUMERICAS = {"quantidade", "valor", "quantidade_destino", "percentual_custo", "valor_corrente"}


def _decimal(texto: str, campo: str, linha: int) -> Decimal | None:
    if not texto:
        return None
    try:
        valor = Decimal(texto)
    except InvalidOperation as erro:
        raise ImportacaoError(f"eventos linha {linha}: {campo} inválido ({texto!r})") from erro
    if not valor.is_finite() or valor < 0:  # o sentido vem da coluna C/V, nunca do sinal
        raise ImportacaoError(f"eventos linha {linha}: {campo} deve ser número >= 0 ({texto!r})")
    return valor


EXIGE_QUANTIDADE = {TipoEvento.DESDOBRAMENTO, TipoEvento.GRUPAMENTO, TipoEvento.BONIFICACAO,
                    TipoEvento.SUBSCRICAO, TipoEvento.CUSTODIA_INICIAL,
                    TipoEvento.OFERTA_PUBLICA, TipoEvento.DIREITO_RECEBIDO}  # fmt: skip


def _chave(linha: dict[str, str], n: int) -> str:
    """Hash dos campos que definem o evento, normalizados ("100" = "100.0"; tipo e C/V em
    maiúsculas); observação e origem não entram. Os campos opcionais só entram preenchidos (a
    chave dos eventos antigos não muda)."""
    nomes = [*COLUNAS[:10], *(c for c in EXTRAS if linha.get(c, ""))]
    campos: list[str] = []
    for nome in nomes:
        campo = linha.get(nome, "").upper()
        if nome in NUMERICAS:
            numero = _decimal(campo, nome, n)
            campo = "" if numero is None else format(numero.normalize(), "f")
        campos.append(campo if nome in COLUNAS else f"{nome}={campo}")
    return hashlib.sha256("|".join(campos).encode()).hexdigest()


def _chaves(linha: dict[str, str], n: int) -> tuple[str, str]:
    """A chave da linha e a da mesma linha com o valor zero escrito do outro jeito ("" ou "0"):
    o mesmo evento lançado na tela e no CSV não duplica."""
    valor = _decimal(linha.get("valor", ""), "valor", n)
    outra = "0" if not linha.get("valor", "") else ("" if valor == 0 else linha["valor"])
    return _chave(linha, n), _chave({**linha, "valor": outra}, n)


# ------------------------------------------------------------------------ tela 7 (Eventos)
NOTA_DATA_EX = ("a B3 informa a data do crédito: use a data ex (primeiro pregão sem o direito), "
                "como decidido para os eventos")  # fmt: skip


SEM_CV = frozenset({TipoEvento.TRANSFERENCIA, *TIPOS_COM_DESTINO, TipoEvento.LEILAO_FRACAO,
                    TipoEvento.DIREITO_NAO_EXERCIDO})  # fmt: skip


# direito não exercido sem quantidade baixa todo o saldo: a linha "não exercido" da B3 vem sem ela
# no dividendo em bens a quantidade é só a do bem recebido (a de destino)
QUANTIDADE_OBRIGATORIA = (EXIGE_QUANTIDADE | {*TIPOS_COM_DESTINO, TipoEvento.LEILAO_FRACAO}) - {
    TipoEvento.DIVIDENDO_EM_BENS
}


AVISO_INCORPORACAO_ACOES = (
    "incorporação de ações (Lei 6.404 art. 252): o COIR mantém o custo (posição do Judiciário); a "
    "Receita trata como alienação com ganho de capital (SC Cosit 224/2014). Avalie com o seu "
    "contador.")  # fmt: skip


@dataclass(frozen=True)
class DadosEvento:
    """O que a tela grava (ou propõe a partir da movimentação da B3)."""

    tipo: TipoEvento
    data: date
    conta: int | None
    ticker: str
    cv: str | None = None
    quantidade: Decimal | None = None
    valor: Decimal = Decimal(0)
    conta_destino: int | None = None
    ticker_destino: str | None = None
    quantidade_destino: Decimal | None = None
    natureza: NaturezaEvento | None = None
    percentual_custo: Decimal | None = None
    observacao: str = ""
    origem: str = "tela"
    valor_corrente: Decimal | None = None  # dividendo em bens: valor de mercado na entrega


@dataclass(frozen=True)
class EventoLancado:
    id: int
    conta_nome: str  # "XP 123456"
    conta_destino_nome: str | None
    dados: DadosEvento


@dataclass(frozen=True)
class Proposta:
    """Linha da movimentação da B3 que sugere um evento. ``situacao``: pendente, lançado (há
    evento do ativo perto da data, ou troca de ticker em data/) ou ignorada."""

    chave: str
    data: date
    movimentacao: str
    codigo: str
    quantidade: Decimal
    valor: Decimal | None
    corretora: str | None
    situacao: str
    dados: DadosEvento
    nota: str
    segura: bool = False  # lançada sozinha, sem pedir nada ao usuário (ver _segura)
    # depois da última nota: sem saber se o titular ainda tinha o ativo, fica fora da conferência
    depois_da_ultima_nota: bool = False
    custo_por_acao: Decimal | None = None  # bonificação: como veio do aviso aos acionistas
    referencia: str = ""  # de onde veio o custo (documento na CVM)


def _nome_conta(conta: Conta) -> str:
    return f"{conta.corretora.nome} {conta.numero}"


def listar(banco: Path, titular: int) -> list[EventoLancado]:
    """Eventos do titular, mais recentes primeiro."""
    with conectar(banco) as fabrica, fabrica() as s:
        tickers = Tickers(s)
        saida: list[EventoLancado] = []
        for e in s.scalars(
            select(EventoCorporativo).join(Conta, EventoCorporativo.conta_id == Conta.id)
            .where(Conta.titular_id == titular)
            .options(selectinload(EventoCorporativo.conta).selectinload(Conta.corretora),
                     selectinload(EventoCorporativo.conta_destino).selectinload(Conta.corretora))
            .order_by(EventoCorporativo.data.desc(), EventoCorporativo.id.desc())
        ):  # fmt: skip
            destino = (None if e.ativo_destino_id is None
                       else tickers.ticker(e.ativo_destino_id, e.data))  # fmt: skip
            dados = DadosEvento(
                tipo=TipoEvento(e.tipo), data=e.data, conta=e.conta_id,
                ticker=tickers.ticker(e.ativo_id, e.data), cv=e.cv, quantidade=e.quantidade,
                valor=e.valor, conta_destino=e.conta_destino_id, ticker_destino=destino,
                quantidade_destino=e.quantidade_destino,
                natureza=None if e.natureza is None else NaturezaEvento(e.natureza),
                percentual_custo=e.percentual_custo, observacao=e.observacao, origem=e.origem,
                valor_corrente=e.valor_corrente,
            )  # fmt: skip
            destino_nome = None if e.conta_destino is None else _nome_conta(e.conta_destino)
            saida.append(EventoLancado(e.id, _nome_conta(e.conta), destino_nome, dados))
        return saida


def _validar(d: DadosEvento) -> None:
    def erro(motivo: str) -> ImportacaoError:
        return ImportacaoError(f"{d.tipo.value.lower().replace('_', ' ')}: {motivo}")

    if not d.ticker.strip():
        raise erro("informe o ativo")
    if d.conta is None:
        raise erro("informe a conta")
    if d.tipo in SEM_CV:
        if d.cv is not None:
            raise erro("este tipo não tem C/V")
    elif d.cv not in ("C", "V"):
        raise erro("informe o C/V: C soma, V subtrai")
    for nome, numero in (("quantidade", d.quantidade), ("valor", d.valor),
                         ("quantidade recebida", d.quantidade_destino),
                         ("valor de mercado", d.valor_corrente)):  # fmt: skip
        if numero is not None and (not numero.is_finite() or numero < 0):
            raise erro(f"{nome} deve ser um número >= 0")
    if d.tipo in QUANTIDADE_OBRIGATORIA and not d.quantidade:
        raise erro("informe a quantidade")
    if d.tipo in (TipoEvento.REDUCAO_CAPITAL, TipoEvento.LEILAO_FRACAO) and not d.valor:
        raise erro("informe o valor")
    if d.tipo is TipoEvento.OFERTA_PUBLICA:
        if d.cv != "C":
            raise erro("a oferta pública é uma compra (C)")
        if not d.valor:
            raise erro("informe o valor pago")
    if d.tipo is TipoEvento.DIREITO_RECEBIDO and (d.cv != "C" or d.valor):
        raise erro("o direito recebido entra com custo zero; o direito comprado vem na nota de "
                   "corretagem")  # fmt: skip
    if d.tipo is TipoEvento.DIVIDENDO_EM_BENS:
        if not d.valor:
            raise erro("informe o valor atribuído aos bens no aviso aos acionistas (é o "
                       "dividendo, rendimento isento)")  # fmt: skip
        if not d.valor_corrente:
            # critério do COIR (16/11/2026): o custo é o menor entre o atribuído e o de mercado
            raise erro("informe o valor de mercado dos bens na data da entrega (quantidade vezes "
                       "a cotação de fechamento do dia)")  # fmt: skip
        if d.quantidade:
            raise erro("informe só a quantidade recebida")
    elif d.valor_corrente is not None:
        raise erro("valor de mercado na entrega só vale no dividendo em bens")
    if d.tipo is TipoEvento.EXERCICIO_DIREITO and not d.valor:
        raise erro("informe o valor pago na subscrição (quantidade vezes o preço de subscrição)")
    if d.tipo is TipoEvento.CUSTODIA_INICIAL:
        # critérios do COIR (25/10/2026): só em 31/12 e só posição comprada
        if (d.data.month, d.data.day) != (12, 31):
            raise erro("a custódia inicial é a posição em 31/12 (a do Bens e Direitos da última "
                       "declaração)")  # fmt: skip
        if d.cv != "C":
            raise erro("lance a posição comprada (C); posição vendida pede as notas")
        if not d.valor:
            raise erro("informe o custo de aquisição (o valor do Bens e Direitos)")
    if d.tipo is TipoEvento.TRANSFERENCIA and (d.conta_destino is None
                                               or d.conta_destino == d.conta):  # fmt: skip
        raise erro("informe a conta de destino (outra conta do titular)")
    if d.tipo in TIPOS_COM_DESTINO:
        if not d.ticker_destino or d.ticker_destino.strip().upper() == d.ticker.strip().upper():
            raise erro("informe o ativo recebido (diferente do de origem)")
        if not d.quantidade_destino:
            raise erro("informe a quantidade recebida")
    if d.tipo is TipoEvento.INCORPORACAO and d.natureza is None:
        raise erro("informe a natureza (incorporação, fusão, conversão, ajuste de série...)")
    if d.tipo not in TIPOS_COM_DESTINO and (d.ticker_destino or d.quantidade_destino):
        raise erro("ativo e quantidade recebidos só valem em incorporação, cisão, exercício de "
                   "direito de subscrição e dividendo em bens")  # fmt: skip
    outra_classe = (d.tipo in (TipoEvento.BONIFICACAO, TipoEvento.SUBSCRICAO)
                    and d.natureza is NaturezaEvento.OUTRA_CLASSE)  # fmt: skip
    if d.tipo is not TipoEvento.INCORPORACAO and d.natureza is not None and not outra_classe:
        raise erro("natureza só vale na incorporação")
    if d.tipo is not TipoEvento.CISAO and d.percentual_custo is not None:
        raise erro("percentual do custo só vale na cisão")
    if d.tipo is TipoEvento.CISAO and not (d.percentual_custo and 0 < d.percentual_custo < 1):
        # critério do COIR (P4, 19/10/2026): sem o percentual divulgado, não calcula
        raise erro("informe o percentual do custo que vai para o ativo recebido (aviso aos "
                   "acionistas da companhia)")  # fmt: skip


def _campos_chave(conta: Conta, destino: Conta | None, d: DadosEvento) -> dict[str, str]:
    """Mesmos campos do CSV (a mesma linha no CSV e na tela tem a mesma chave)."""

    def numero(x: Decimal | None) -> str:
        return "" if x is None else format(x, "f")

    return {
        "data": d.data.isoformat(), "corretora": conta.corretora.nome, "conta": conta.numero,
        "ticker": d.ticker, "tipo": d.tipo.value, "cv": d.cv or "",
        "quantidade": numero(d.quantidade), "valor": numero(d.valor),
        "corretora_destino": "" if destino is None else destino.corretora.nome,
        "conta_destino": "" if destino is None else destino.numero,
        "ticker_destino": d.ticker_destino or "",
        "quantidade_destino": numero(d.quantidade_destino),
        "natureza": "" if d.natureza is None else d.natureza.value,
        "percentual_custo": numero(d.percentual_custo),
        "valor_corrente": numero(d.valor_corrente),
    }  # fmt: skip


def _preencher(
    s: Session, registro: EventoCorporativo, d: DadosEvento, chave: str, ativo: int
) -> None:  # fmt: skip
    if d.conta is None:
        raise ImportacaoError("informe a conta")
    registro.chave, registro.data, registro.conta_id, registro.ativo_id = (
        chave, d.data, d.conta, ativo)  # fmt: skip
    registro.tipo, registro.cv, registro.quantidade, registro.valor = (
        d.tipo.value, d.cv, d.quantidade, d.valor)  # fmt: skip
    registro.conta_destino_id = d.conta_destino
    # ativo recebido que nunca foi negociado (ex.: AMOB3, série com código novo) é cadastrado
    registro.ativo_destino_id = (
        None if d.ticker_destino is None else repo.Ativos(s).ativo_id(d.ticker_destino, d.data)
    )
    registro.quantidade_destino, registro.percentual_custo = (
        d.quantidade_destino, d.percentual_custo)  # fmt: skip
    registro.valor_corrente = d.valor_corrente
    registro.natureza = None if d.natureza is None else d.natureza.value
    registro.observacao, registro.origem = d.observacao, d.origem


JANELA = 10  # dias entre a data da B3 (crédito) e a do evento lançado (data ex)


def _ativos_do_ticker(s: Session, ticker: str) -> set[int]:
    return set(s.scalars(select(TickerHistorico.ativo_id).where(TickerHistorico.ticker == ticker)))


def _ja_lancado(
    s: Session, refs: Referencias, ticker: str, dia: date,
    lancados: list[tuple[int, int | None, date]], troca_conta: bool = True,
) -> bool:  # fmt: skip
    """Há evento do ativo (de origem ou recebido) a até JANELA dias, ou troca de código em
    data/trocas_ticker.csv (ex.: incorporação JSLG3 -> SIMH3 tratada como troca). A repartição do
    custo de uma incorporação com dinheiro (``troca_conta=False``) vem junto com a troca."""
    perto = timedelta(days=JANELA)
    if troca_conta and any(t.ticker == ticker and abs(t.ate - dia) <= perto for t in refs.trocas):
        return True
    ativo = repo.ativo_existente(s, ticker, dia)
    if ativo is None:  # código atual de ativo que se chamava outra coisa na data
        candidatos = _ativos_do_ticker(s, ticker)
        ativo = next(iter(candidatos)) if len(candidatos) == 1 else None
    return ativo is not None and any(
        ativo in (origem, destino) and abs(quando - dia) <= perto
        for origem, destino, quando in lancados
    )


def ignorar(banco: Path, titular: int, chave: str, agora: datetime, sim: bool = True) -> None:
    """Marca (ou desmarca) a proposta como ignorada."""
    with conectar(banco) as fabrica, fabrica.begin() as s:
        registro = s.scalar(select(PropostaIgnorada).where(PropostaIgnorada.chave == chave))
        if sim and registro is None:
            s.add(PropostaIgnorada(titular_id=titular, chave=chave, ignorada_em=agora))
        elif not sim and registro is not None:
            s.delete(registro)


def ignoradas(banco: Path, titular: int) -> set[str]:
    with conectar(banco) as fabrica, fabrica() as s:
        return set(s.scalars(select(PropostaIgnorada.chave)
                             .where(PropostaIgnorada.titular_id == titular)))  # fmt: skip


def contas_do_titular(banco: Path, titular: int) -> dict[int, str]:
    """Contas do titular para escolher na tela: id -> "XP 123456"."""
    with conectar(banco) as fabrica, fabrica() as s:
        contas = s.scalars(select(Conta).where(Conta.titular_id == titular)
                           .options(selectinload(Conta.corretora)))  # fmt: skip
        return dict(sorted(((c.id, _nome_conta(c)) for c in contas), key=lambda x: x[1]))
