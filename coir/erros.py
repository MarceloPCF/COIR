"""Hierarquia de erros do COIR.

Toda falha de domínio usa uma classe própria (CLAUDE.md, regra 5).
"""


class CoirError(Exception):
    """Base de todos os erros do COIR."""


class ConfiguracaoError(CoirError):
    """Configuração ausente ou inválida (ex.: pasta de dados inexistente)."""
