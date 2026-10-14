"""Esquema do banco (SQLAlchemy 2). Mudanças só por migração Alembic (regra 11).

Cadastros: titular, corretora, conta, ativo, ticker_historico, serie_opcao.
Fatos (imutáveis, regra 7): arquivo_importado, nota_corretagem, nota_custo, operacao.

O ativo tem chave interna; o ISIN é único e opcional (a nota SINACOR não o traz; o pipeline de
referências da F7 o preenche). O mesmo ticker pode ser de ativos diferentes em épocas diferentes:
a resolução é (ticker, data) -> ativo por ``ticker_historico``.
"""

from datetime import date, datetime
from decimal import Decimal

from sqlalchemy import ForeignKey, Index, MetaData, String, Text, UniqueConstraint
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship

from coir.db.tipos import DecimalTexto

CONVENCAO_NOMES = {
    "ix": "ix_%(column_0_label)s",
    "uq": "uq_%(table_name)s_%(column_0_name)s",
    "fk": "fk_%(table_name)s_%(column_0_name)s_%(referred_table_name)s",
    "pk": "pk_%(table_name)s",
}


class Base(DeclarativeBase):
    metadata = MetaData(naming_convention=CONVENCAO_NOMES)
    type_annotation_map = {Decimal: DecimalTexto}  # noqa: RUF012


class Titular(Base):
    __tablename__ = "titular"

    id: Mapped[int] = mapped_column(primary_key=True)
    cpf: Mapped[str] = mapped_column(String(11), unique=True)  # só dígitos
    nome: Mapped[str] = mapped_column(Text)


class Corretora(Base):
    __tablename__ = "corretora"

    id: Mapped[int] = mapped_column(primary_key=True)
    nome: Mapped[str] = mapped_column(String(40), unique=True)  # nome do perfil (AGORA, XP, BTG...)


class Conta(Base):
    __tablename__ = "conta"
    __table_args__ = (UniqueConstraint("corretora_id", "numero"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    corretora_id: Mapped[int] = mapped_column(ForeignKey("corretora.id"))
    numero: Mapped[str] = mapped_column(String(40))  # como na nota normalizada
    titular_id: Mapped[int] = mapped_column(ForeignKey("titular.id"))

    corretora: Mapped[Corretora] = relationship()
    titular: Mapped[Titular] = relationship()


class Ativo(Base):
    __tablename__ = "ativo"

    id: Mapped[int] = mapped_column(primary_key=True)
    isin: Mapped[str | None] = mapped_column(String(12), unique=True)


class TickerHistorico(Base):
    """Ticker de um ativo numa vigência (início/fim inclusivos; vazio = em aberto).

    No SQLite NULL é distinto de NULL: o índice único não impede duas vigências em aberto do
    mesmo ticker; a sobreposição é verificada em ``repositorio.Ativos`` (falha alto)."""

    __tablename__ = "ticker_historico"
    __table_args__ = (
        UniqueConstraint("ticker", "inicio"),
        Index("ix_ticker_historico_ticker", "ticker"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    ativo_id: Mapped[int] = mapped_column(ForeignKey("ativo.id"))
    ticker: Mapped[str] = mapped_column(String(20))
    inicio: Mapped[date | None]
    fim: Mapped[date | None]


class SerieOpcao(Base):
    """Classe do ativo-objeto vista na abertura da série (ex.: 'BBASN380 ON' -> ON)."""

    __tablename__ = "serie_opcao"

    serie: Mapped[str] = mapped_column(String(20), primary_key=True)
    classe: Mapped[str] = mapped_column(String(5))


class ArquivoImportado(Base):
    """PDF lido. ``completo``: todas as notas entraram; só então a reimportação o ignora."""

    __tablename__ = "arquivo_importado"

    id: Mapped[int] = mapped_column(primary_key=True)
    sha256: Mapped[str] = mapped_column(String(64), unique=True)
    nome: Mapped[str] = mapped_column(Text)
    tamanho: Mapped[int]
    importado_em: Mapped[datetime]
    completo: Mapped[bool]


class NotaCorretagem(Base):
    """A nota como impressa (``json_nota``: ``NotaExtraida`` completa) e os totais para consulta."""

    __tablename__ = "nota_corretagem"
    __table_args__ = (UniqueConstraint("conta_id", "numero", "data_pregao", "segmento"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    arquivo_id: Mapped[int] = mapped_column(ForeignKey("arquivo_importado.id"))
    conta_id: Mapped[int] = mapped_column(ForeignKey("conta.id"))
    numero: Mapped[str] = mapped_column(String(40))
    data_pregao: Mapped[date]
    segmento: Mapped[str] = mapped_column(String(10))
    paginas: Mapped[str] = mapped_column(Text)  # "3,4"
    hash_conteudo: Mapped[str] = mapped_column(String(64))
    json_nota: Mapped[str] = mapped_column(Text)
    # Bovespa: "Valor das operações" (sem sinal); BM&F: "Valor dos negócios" (crédito +, débito -)
    valor_operacoes: Mapped[Decimal]
    custos_total: Mapped[Decimal]  # sem IRRF (BM&F: inclui impostos)
    irrf: Mapped[Decimal]  # I.R.R.F. s/ operações impresso (0,005%)
    irrf_day_trade: Mapped[Decimal]  # IRRF de day trade impresso (1%); não somar em irrf
    liquido: Mapped[Decimal]  # crédito positivo, débito negativo
    data_liquidacao: Mapped[date | None]

    arquivo: Mapped[ArquivoImportado] = relationship()
    conta: Mapped[Conta] = relationship()
    custos: Mapped[list["NotaCusto"]] = relationship(order_by="NotaCusto.ordem")
    operacoes: Mapped[list["Operacao"]] = relationship(order_by="Operacao.ordem")


class NotaCusto(Base):
    __tablename__ = "nota_custo"

    id: Mapped[int] = mapped_column(primary_key=True)
    nota_id: Mapped[int] = mapped_column(ForeignKey("nota_corretagem.id"), index=True)
    ordem: Mapped[int]
    rotulo: Mapped[str] = mapped_column(Text)
    categoria: Mapped[str] = mapped_column(String(20))
    grupo: Mapped[str] = mapped_column(String(10))
    valor: Mapped[Decimal]


class Operacao(Base):
    """Operação normalizada da nota (livro imutável). ``papel``: ticker na data (após troca);
    ``ticker_impresso``: antes da troca; ``ativo_id``: ativo resolvido."""

    __tablename__ = "operacao"

    id: Mapped[int] = mapped_column(primary_key=True)
    nota_id: Mapped[int] = mapped_column(ForeignKey("nota_corretagem.id"), index=True)
    ordem: Mapped[int]
    categoria: Mapped[str] = mapped_column(String(10))
    cv: Mapped[str] = mapped_column(String(1))
    papel: Mapped[str] = mapped_column(String(20))
    ticker_impresso: Mapped[str] = mapped_column(String(20))
    ativo_id: Mapped[int] = mapped_column(ForeignKey("ativo.id"), index=True)
    mercado: Mapped[str] = mapped_column(String(20))
    preco: Mapped[Decimal]
    quantidade: Mapped[Decimal]
    total: Mapped[Decimal]
    ativo_objeto: Mapped[str | None] = mapped_column(String(20))
    prazo_termo_dias: Mapped[int | None]
