"""Configuração do COIR.

A pasta de dados reais do usuário (notas, relatórios B3, gabarito, banco) fica FORA do repositório.
O caminho vem da variável de ambiente ``COIR_DADOS``; sem ela, usa ``~/COIR-dados``.
"""

import logging
import os
from collections.abc import Mapping
from pathlib import Path

from coir.erros import ConfiguracaoError

VAR_DADOS = "COIR_DADOS"
PASTA_PADRAO = "COIR-dados"

log = logging.getLogger(__name__)


def dados_dir(*, exigir: bool = False, ambiente: Mapping[str, str] | None = None) -> Path:
    """Retorna a pasta de dados do usuário.

    Args:
        exigir: se True, levanta ``ConfiguracaoError`` quando a pasta não existe.
        ambiente: variáveis de ambiente a consultar (padrão ``os.environ``); útil em testes.
    """
    env = os.environ if ambiente is None else ambiente
    valor = env.get(VAR_DADOS, "").strip()
    if valor:
        pasta = Path(valor).expanduser()
        origem = VAR_DADOS
    else:
        pasta = Path.home() / PASTA_PADRAO
        origem = "padrão"
    log.debug("pasta de dados (%s): %s", origem, pasta)
    if exigir and not pasta.is_dir():
        raise ConfiguracaoError(
            f"Pasta de dados não encontrada: {pasta} (origem: {origem}). "
            f"Defina a variável de ambiente {VAR_DADOS} com o caminho correto."
        )
    return pasta
