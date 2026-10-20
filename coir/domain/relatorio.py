"""Entrada dos relatórios (F5): o que o serviço tira do banco e do motor para um titular e um ano.
Os relatórios (`coir.reports`) partem só disto, sem tocar no banco nem no motor."""

from datetime import date, datetime

from pydantic import BaseModel, ConfigDict

from coir.domain.fiscal import Apuracao, PosicaoAtivo, Visao
from coir.domain.irpf import MapeamentoIrpf
from coir.domain.nota import NotaExtraida, OperacaoExtraida
from coir.domain.proventos import Provento


class NotaDoLivro(BaseModel):
    model_config = ConfigDict(frozen=True, extra="forbid")

    nota: NotaExtraida  # como extraída (sem CPF nem nome do cliente)
    operacoes: tuple[OperacaoExtraida, ...]


class InfoAtivo(BaseModel):
    """Cadastro público de um ticker para a declaração (pacote da B3 e tabelas curadas)."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    emissor: str  # 4 letras (do cadastro de instrumentos; sem ele, a raiz do ticker)
    cnpj: str | None  # 14 dígitos; None se o emissor não está no pacote nem em emissores.csv
    razao_social: str
    especificacao: str = ""  # "ON NM", "PN N1", "UNT N2", "CI", "DRN"
    tipo_b3: str = ""  # COMPANHIA, FII, FIAGRO, ETF, ETF-RF...
    objeto: str | None = None  # opção: ticker do ativo-objeto


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
    carteira_fiscal_anterior: tuple[PosicaoAtivo, ...] = ()  # em 31/12 do ano anterior
    contas: dict[int, str]  # id da conta -> número impresso (carteira por conta)
    notas: tuple[NotaDoLivro, ...]  # do titular, no ano
    proventos: tuple[Provento, ...] = ()  # do titular (todos os anos: o relatório filtra)
    avisos_proventos: tuple[str, ...] = ()
    # declaração anual (F8): mapeamento do exercício, cadastro dos tickers e custódia por
    # corretora em 31/12 (só para a discriminação dos bens; o custo é o fiscal, por titular)
    irpf: MapeamentoIrpf | None = None  # None: sem data/irpf/<exercício>.yaml
    info_ativos: dict[str, InfoAtivo] = {}
    custodia: tuple[PosicaoAtivo, ...] = ()  # por corretora, em 31/12 do ano (ou em `ate`)
    custodia_anterior: tuple[PosicaoAtivo, ...] = ()  # por corretora, em 31/12 do ano anterior
