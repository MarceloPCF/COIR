"""Hierarquia de erros do COIR.

Toda falha de domínio usa uma classe própria (CLAUDE.md, regra 5).
"""

from datetime import date

from coir.auditoria import erro_de_dominio


class CoirError(Exception):
    """Base de todos os erros do COIR. Todo erro nasce registrado na auditoria, porque a tela só
    mostra a mensagem. Ficam de fora os do caminho normal, registrados pelo resultado: a nota
    rejeitada (``ExtracaoError``), o ativo isolado (``MotorError``) e a ordem que a tela Simulação
    recusa (``SimulacaoError``: digitação do usuário, e o motor não escreve em disco)."""

    def __init__(self, *args: object) -> None:
        super().__init__(*args)
        if not isinstance(self, ExtracaoError | MotorError | SimulacaoError):
            erro_de_dominio(self)


class ConfiguracaoError(CoirError):
    """Configuração ausente ou inválida (ex.: pasta de dados inexistente)."""


class GabaritoError(CoirError):
    """Planilha-gabarito com estrutura ou conteúdo inesperado."""


class ExtracaoError(CoirError):
    """Texto da nota fora do layout esperado (campo ausente, número ilegível, página sem perfil)."""


class ConciliacaoError(ExtracaoError):
    """Os valores lidos da nota não fecham entre si (regra 4): a nota não entra."""

    def __init__(self, campo: str, esperado: object, lido: object, contexto: str = "") -> None:
        self.campo, self.esperado, self.lido = campo, esperado, lido
        super().__init__(f"{contexto}{campo}: esperado {esperado}, lido {lido}")


class ReferenciaError(ExtracaoError):
    """Ativo, opção ou contrato sem cadastro nas tabelas de referência (data/)."""


class NomeAmbiguoError(ReferenciaError):
    """O nome de pregão impresso na nota valia para mais de um código na data, e nada desempatou
    (planilha de negociação da B3, escolha guardada). ``nome``: o nome compacto que ficou em
    dúvida; ``codigos``: os candidatos."""

    def __init__(
        self, mensagem: str, especificacao: str, nome: str, codigos: tuple[str, ...]
    ) -> None:
        super().__init__(mensagem)
        self.especificacao, self.nome, self.codigos = especificacao, nome, codigos


class BancoError(CoirError):
    """Banco de dados ausente, desatualizado ou com conteúdo inesperado."""


class ImportacaoError(CoirError):
    """Lote de importação que não pode ser gravado (titular, conta ou arquivos inconsistentes)."""


class MotorError(CoirError):
    """Livro que o motor fiscal não consegue apurar sem palpite (ex.: exercício sem a opção em
    carteira, venda a termo, custo negativo).

    ``data``: dia do livro em que o motor parou (None quando falha antes de percorrer os dias).
    Antes dela o livro fechou; quem só precisa das posições até ali segue sem o erro."""

    def __init__(self, mensagem: str, data: date | None = None) -> None:
        super().__init__(mensagem)
        self.data = data


class AnonimizacaoError(CoirError):
    """Nota cuja cópia anonimizada não passou num dos portões (fidelidade ou privacidade): a
    cópia não é gravada."""


class ExportacaoError(CoirError):
    """Relatório inconsistente (linha com colunas a mais ou a menos) ou destino inválido."""


class AtualizacaoError(CoirError):
    """Pacote de tabelas de referência que não pode ser baixado ou usado (sem rede, hash que não
    confere, formato de outra versão): nada é alterado."""


class SimulacaoError(CoirError):
    """Ordem da tela Simulação que o motor não pode apurar como digitada (data sem pregão, venda
    acima da posição, ativo sem tipo ou isolado, mês fora do calendário)."""
