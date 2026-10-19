"""Entrada dos relatórios (F5): o que o serviço tira do banco e do motor para um titular e um ano.
Os relatórios (`coir.reports`) partem só disto, sem tocar no banco nem no motor."""

from datetime import date, datetime

from pydantic import BaseModel, ConfigDict

from coir.domain.fiscal import Apuracao, PosicaoAtivo, Visao
from coir.domain.nota import NotaExtraida, OperacaoExtraida
from coir.domain.proventos import Provento


class NotaDoLivro(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")

    nota: NotaExtraida  # como extraída (sem CPF nem nome do cliente)
    operacoes: tuple[OperacaoExtraida, ...]


class DadosRelatorio(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")

    titular: int
    nome: str
    cpf: str  # 11 dígitos (cabeçalho dos relatórios; nunca no nome do arquivo)
    ano: int
    gerado_em: datetime
    codigo_darf: str
    apuracao: Apuracao  # do livro inteiro (o saldo de prejuízo vem de anos anteriores)
    ate: date  # a apuração vai até aqui: 31/12 ou a data de geração no ano em curso
    data_carteira: date
    visao: Visao
    carteira: tuple[PosicaoAtivo, ...]
    carteira_fiscal: tuple[PosicaoAtivo, ...]  # sempre por titular (Bens e Direitos)
    contas: dict[int, str]  # id da conta -> número impresso (carteira por conta)
    notas: tuple[NotaDoLivro, ...]  # do titular, no ano
    proventos: tuple[Provento, ...] = ()  # do titular (todos os anos: o relatório filtra)
    avisos_proventos: tuple[str, ...] = ()
