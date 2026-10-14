"""Leitura e gravação dos cadastros e fatos. Sem commit: a transação é de quem chama (regra 8)."""

import hashlib
from collections import defaultdict
from collections.abc import Iterable
from datetime import date, timedelta

from sqlalchemy import func, select
from sqlalchemy.orm import Session, selectinload

from coir.db.modelos import (
    ArquivoImportado,
    Ativo,
    Conta,
    Corretora,
    NotaCorretagem,
    NotaCusto,
    Operacao,
    SerieOpcao,
    TickerHistorico,
    Titular,
)
from coir.domain.cpf import cpf_valido, so_digitos
from coir.domain.nota import NotaExtraida, OperacaoExtraida, Segmento
from coir.domain.referencias import TrocaTicker
from coir.erros import BancoError, ImportacaoError

UM_DIA = timedelta(days=1)


# ------------------------------------------------------------------------------------- titulares
def adicionar_titular(sessao: Session, cpf: str, nome: str) -> Titular:
    digitos = so_digitos(cpf)
    if not cpf_valido(digitos):
        raise ImportacaoError("CPF inválido")
    if titular_por_cpf(sessao, digitos) is not None:
        raise ImportacaoError("titular já cadastrado com esse CPF")
    titular = Titular(cpf=digitos, nome=nome.strip())
    sessao.add(titular)
    sessao.flush()
    return titular


def titular_por_cpf(sessao: Session, cpf: str) -> Titular | None:
    return sessao.scalar(select(Titular).where(Titular.cpf == so_digitos(cpf)))


def titulares(sessao: Session) -> list[Titular]:
    return list(sessao.scalars(select(Titular).order_by(Titular.nome, Titular.id)))


# ---------------------------------------------------------------------------- corretoras e contas
class Contas:
    """Corretoras e contas do lote, com cache. Conta nova é do titular do lote; conta existente
    de outro titular é erro (nunca mistura CPFs)."""

    def __init__(self, sessao: Session, titular: Titular | None) -> None:
        self._sessao = sessao
        self._titular = titular
        self._corretoras: dict[str, Corretora] = {}
        self._contas: dict[tuple[str, str], Conta] = {}

    def corretora(self, nome: str) -> Corretora:
        if nome not in self._corretoras:
            existente = self._sessao.scalar(select(Corretora).where(Corretora.nome == nome))
            if existente is None:
                existente = Corretora(nome=nome)
                self._sessao.add(existente)
                self._sessao.flush()
            self._corretoras[nome] = existente
        return self._corretoras[nome]

    def conta(self, corretora: str, numero: str) -> Conta:
        chave = (corretora, numero)
        if chave in self._contas:
            return self._contas[chave]
        c = self.corretora(corretora)
        conta = self._sessao.scalar(
            select(Conta).where(Conta.corretora_id == c.id, Conta.numero == numero)
        )
        if conta is None:
            if self._titular is None:
                raise ImportacaoError(
                    f"conta nova ({corretora}) e nenhum titular definido: informe --titular"
                )
            conta = Conta(corretora=c, numero=numero, titular=self._titular)
            self._sessao.add(conta)
            self._sessao.flush()
        elif self._titular is not None and conta.titular_id != self._titular.id:
            raise ImportacaoError(
                f"conta da {corretora} já pertence a outro titular; importe esse arquivo no "
                "titular dono da conta"
            )
        self._contas[chave] = conta
        return conta


# ---------------------------------------------------------------------------------------- ativos
class Ativos:
    """Resolução (ticker, data) -> ativo pelo histórico de tickers, com cache do lote.

    Ticker sem vigência que cubra a data cria um ativo novo, com vigência a partir do dia seguinte
    ao último fim conhecido desse ticker (ou em aberto, se o ticker é inédito).
    """

    def __init__(self, sessao: Session) -> None:
        self._sessao = sessao
        self._por_ticker: dict[str, list[TickerHistorico]] = defaultdict(list)
        for linha in sessao.scalars(select(TickerHistorico)):
            self._por_ticker[linha.ticker].append(linha)

    def _novo_ativo(self) -> Ativo:
        ativo = Ativo()
        self._sessao.add(ativo)
        self._sessao.flush()
        return ativo

    def _adicionar(self, ativo: Ativo, ticker: str, inicio: date | None, fim: date | None) -> None:
        self._adicionar_id(ativo.id, ticker, inicio, fim)

    def _cobrindo(self, ticker: str, dia: date) -> TickerHistorico | None:
        linhas = [
            x
            for x in self._por_ticker.get(ticker, [])
            if (x.inicio is None or x.inicio <= dia) and (x.fim is None or dia <= x.fim)
        ]
        if len(linhas) > 1:
            raise BancoError(f"ticker {ticker} com vigências sobrepostas em {dia}")
        return linhas[0] if linhas else None

    def _lacuna(self, ticker: str, dia: date) -> tuple[date | None, date | None]:
        """Vigência livre do ticker em volta de ``dia``: do fim anterior ao início seguinte."""
        linhas = self._por_ticker.get(ticker, [])
        fins = [x.fim for x in linhas if x.fim is not None and x.fim < dia]
        inicios = [x.inicio for x in linhas if x.inicio is not None and x.inicio > dia]
        return (max(fins) + UM_DIA if fins else None, min(inicios) - UM_DIA if inicios else None)

    def ativo_id(self, ticker: str, dia: date) -> int:
        linha = self._cobrindo(ticker, dia)
        if linha is not None:
            return linha.ativo_id
        ativo = self._novo_ativo()
        self._adicionar(ativo, ticker, *self._lacuna(ticker, dia))
        return ativo.id

    def sincronizar_trocas(self, trocas: Iterable[TrocaTicker]) -> bool:
        """Troca de código (data/trocas_ticker.csv): o ticker antigo até ``ate`` e o novo a partir
        do dia seguinte são o mesmo ativo. Idempotente; aceita troca em cadeia (A->B, depois B->C)
        e troca acrescentada depois de já haver operações gravadas (reatribui o ativo delas).
        Devolve se o histórico mudou."""
        mudou = False
        for t in trocas:
            inicio_novo = t.ate + UM_DIA
            antiga = self._cobrindo(t.ticker_impresso, t.ate)
            if antiga is not None and antiga.fim == t.ate:
                continue  # já sincronizada
            nova = self._cobrindo(t.ticker, inicio_novo)
            if (antiga is not None and antiga.fim is not None) or (
                nova is not None and nova.inicio is not None and nova.inicio != inicio_novo
            ):
                raise BancoError(
                    f"troca {t.ticker_impresso}->{t.ticker} em {t.ate} conflita com o histórico "
                    "de tickers (dois tickers trocados para o mesmo código? use evento corporativo)"
                )
            if antiga is not None:
                ativo_id = antiga.ativo_id
                antiga.fim = t.ate
            else:
                ativo_id = nova.ativo_id if nova is not None else self._novo_ativo().id
                inicio, _ = self._lacuna(t.ticker_impresso, t.ate)
                self._adicionar_id(ativo_id, t.ticker_impresso, inicio, t.ate)
            if nova is not None:
                nova.inicio, nova.ativo_id = inicio_novo, ativo_id
            else:
                _, fim = self._lacuna(t.ticker, inicio_novo)
                self._adicionar_id(ativo_id, t.ticker, inicio_novo, fim)
            mudou = True
        self._sessao.flush()
        if mudou:
            self._reatribuir_operacoes()
        return mudou

    def _adicionar_id(
        self, ativo_id: int, ticker: str, inicio: date | None, fim: date | None
    ) -> None:
        linha = TickerHistorico(ativo_id=ativo_id, ticker=ticker, inicio=inicio, fim=fim)
        self._sessao.add(linha)
        self._por_ticker[ticker].append(linha)

    def _reatribuir_operacoes(self) -> None:
        """O ativo da operação é derivado (ticker impresso + data): recalcula após mudar o
        histórico. O livro (o que a nota diz) não muda."""
        linhas = self._sessao.execute(
            select(Operacao, NotaCorretagem.data_pregao).join(NotaCorretagem)
        ).all()
        for op, dia in linhas:
            op.ativo_id = self.ativo_id(op.ticker_impresso, dia)
        self._sessao.flush()


# --------------------------------------------------------------------------------- séries de opção
def series(sessao: Session) -> dict[str, str]:
    return {s.serie: s.classe for s in sessao.scalars(select(SerieOpcao))}


def gravar_series(sessao: Session, novas: dict[str, str]) -> None:
    existentes = series(sessao)
    sessao.add_all(
        SerieOpcao(serie=s, classe=c) for s, c in sorted(novas.items()) if s not in existentes
    )


# ---------------------------------------------------------------------------------- arquivos/notas
def arquivo_por_sha(sessao: Session, sha256: str) -> ArquivoImportado | None:
    return sessao.scalar(select(ArquivoImportado).where(ArquivoImportado.sha256 == sha256))


def hash_conteudo(nota: NotaExtraida) -> str:
    """Hash do que a nota diz (negócios e resumo), sem arquivo/páginas: a mesma nota em outro PDF
    tem o mesmo hash."""
    campos = {"negocios_bovespa", "negocios_bmf", "resumo_bovespa", "resumo_bmf"}
    return hashlib.sha256(nota.model_dump_json(include=campos).encode()).hexdigest()


def nota_por_chave(
    sessao: Session, chave: tuple[str, str, str, date, Segmento]
) -> NotaCorretagem | None:
    corretora, conta, numero, data_pregao, segmento = chave
    return sessao.scalar(
        select(NotaCorretagem)
        .join(Conta)
        .join(Corretora)
        .where(
            Corretora.nome == corretora,
            Conta.numero == conta,
            NotaCorretagem.numero == numero,
            NotaCorretagem.data_pregao == data_pregao,
            NotaCorretagem.segmento == segmento.value,
        )
    )


def gravar_nota(
    sessao: Session,
    nota: NotaExtraida,
    ops: Iterable[OperacaoExtraida],
    arquivo: ArquivoImportado,
    conta: Conta,
    ativos: Ativos,
) -> NotaCorretagem:
    resumo_b, resumo_f = nota.resumo_bovespa, nota.resumo_bmf
    if resumo_b is not None:
        valor, liquido, liquidacao = (
            resumo_b.valor_operacoes,
            resumo_b.liquido,
            resumo_b.data_liquidacao,
        )
        irrf, irrf_dt, custos = resumo_b.irrf, resumo_b.irrf_day_trade, resumo_b.custos
    elif resumo_f is not None:
        valor, liquido, liquidacao = resumo_f.valor_negocios.assinado, resumo_f.liquido, None
        irrf, irrf_dt, custos = resumo_f.irrf, resumo_f.irrf_day_trade, resumo_f.custos
    else:
        raise BancoError(f"nota {nota.numero} sem resumo")
    registro = NotaCorretagem(
        arquivo=arquivo,
        conta=conta,
        numero=nota.numero,
        data_pregao=nota.data_pregao,
        segmento=nota.segmento.value,
        paginas=",".join(map(str, nota.paginas)),
        hash_conteudo=hash_conteudo(nota),
        json_nota=nota.model_dump_json(),
        valor_operacoes=valor,
        custos_total=nota.custos_total,
        irrf=irrf,
        irrf_day_trade=irrf_dt,
        liquido=liquido.assinado,
        data_liquidacao=liquidacao,
        custos=[
            NotaCusto(
                ordem=i,
                rotulo=c.rotulo,
                categoria=str(c.categoria),
                grupo=str(c.grupo),
                valor=c.valor,
            )
            for i, c in enumerate(custos)
        ],
        operacoes=[
            Operacao(
                ordem=i,
                categoria=o.categoria,
                cv=o.cv,
                papel=o.papel,
                ticker_impresso=o.ticker_original,
                ativo_id=ativos.ativo_id(o.ticker_original, o.data),
                mercado=o.mercado,
                preco=o.preco,
                quantidade=o.quantidade,
                total=o.total,
                ativo_objeto=o.ativo_objeto,
                prazo_termo_dias=o.prazo_termo_dias,
            )
            for i, o in enumerate(ops)
        ],
    )
    sessao.add(registro)
    return registro


def _operacao(nota: NotaCorretagem, o: Operacao) -> OperacaoExtraida:
    return OperacaoExtraida.model_validate(
        dict(
            corretora=nota.conta.corretora.nome,
            conta=nota.conta.numero,
            nota=nota.numero,
            data=nota.data_pregao,
            categoria=o.categoria,
            cv=o.cv,
            papel=o.papel,
            mercado=o.mercado,
            preco=o.preco,
            quantidade=o.quantidade,
            total=o.total,
            ativo_objeto=o.ativo_objeto,
            prazo_termo_dias=o.prazo_termo_dias,
            ticker_impresso=o.ticker_impresso if o.ticker_impresso != o.papel else None,
        )
    )


def carregar_notas(sessao: Session) -> list[tuple[NotaExtraida, list[OperacaoExtraida]]]:
    """Notas gravadas (como extraídas) e suas operações, na ordem de pregão e de gravação."""
    notas = sessao.scalars(
        select(NotaCorretagem)
        .options(
            selectinload(NotaCorretagem.operacoes),
            selectinload(NotaCorretagem.conta).selectinload(Conta.corretora),
        )
        .order_by(NotaCorretagem.data_pregao, NotaCorretagem.id)
    )
    return [
        (NotaExtraida.model_validate_json(n.json_nota), [_operacao(n, o) for o in n.operacoes])
        for n in notas
    ]


def contagens(sessao: Session) -> dict[str, int]:
    """Linhas por tabela (para conferir que a reimportação não mudou nada)."""
    tabelas = (Titular, Corretora, Conta, Ativo, TickerHistorico, SerieOpcao, ArquivoImportado,
               NotaCorretagem, NotaCusto, Operacao)  # fmt: skip
    return {
        t.__tablename__: sessao.scalar(select(func.count()).select_from(t)) or 0 for t in tabelas
    }
