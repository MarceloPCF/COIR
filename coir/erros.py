"""Hierarquia de erros do COIR.

Toda falha de domínio usa uma classe própria (CLAUDE.md, regra 5).
"""

from datetime import date


class CoirError(Exception):
    """Base de todos os erros do COIR."""


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


class ExportacaoError(CoirError):
    """Relatório inconsistente (linha com colunas a mais ou a menos) ou destino inválido."""


class AtualizacaoError(CoirError):
    """Pacote de tabelas de referência que não pode ser baixado ou usado (sem rede, hash que não
    confere, formato de outra versão): nada é alterado."""
