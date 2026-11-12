"""Conexão que o navegador abandona: o servidor continua aceitando e o registro não vira erro.

No Windows, o laço de eventos do asyncio (Proactor) tropeça quando o navegador derruba uma conexão
(troca rápida de tela, aba fechada):

1. Se ela cai antes de o servidor terminar de aceitá-la, o accept falha com [WinError 64]
   (ERROR_NETNAME_DELETED) e o asyncio **fecha o socket de escuta**: o COIR para de aceitar
   conexões e a janela fica em "conexão perdida" para sempre. É o python/cpython#93821, corrigido
   no Python 3.14 (PR #161238), não no 3.13. ``laco`` monta o laço com um proactor que, nesses
   erros, volta a aceitar, como a correção do Python; qualquer outro erro segue para o asyncio.
2. Duas mensagens de erro que não dizem nada do COIR: o ``ConnectionResetError`` [WinError 10054]
   de ``_call_connection_lost`` ao fechar a conexão já derrubada, e a tarefa interna do accept
   abandonado ("Task exception was never retrieved"). ``instalar_filtro`` rebaixa só essas duas
   para DEBUG (aparecem com ``--verbose``); os outros erros do asyncio seguem como estão.
"""

import asyncio
import logging
import socket
import sys
from typing import Any

log = logging.getLogger(__name__)

ERROR_NETNAME_DELETED = 64
WSAECONNRESET = 10054
CONEXAO_DERRUBADA = "Exception in callback _ProactorBasePipeTransport._call_connection_lost("
TAREFA_ABANDONADA = "Task exception was never retrieved"
TAREFA_DO_ACCEPT = "coro=<IocpProactor.accept.<locals>.accept_coro()"


def cliente_desistiu(erro: BaseException | None) -> bool:
    """O cliente derrubou a conexão antes de ela ser aceita: os mesmos erros que o Python 3.14
    trata no accept. O 64 chega como OSError comum; 10053, 10054 e 1236 já chegam como
    ConnectionAbortedError ou ConnectionResetError."""
    return isinstance(erro, ConnectionResetError | ConnectionAbortedError) or (
        isinstance(erro, OSError) and getattr(erro, "winerror", None) == ERROR_NETNAME_DELETED
    )


# ------------------------------------------------------------------------------ laço de eventos
if sys.platform == "win32":

    class _Proactor(asyncio.IocpProactor):
        """O accept que não desiste por causa de um cliente: o asyncio só vê a conexão aceita ou um
        erro de verdade (que continua fechando o socket de escuta, como antes)."""

        def accept(self, listener: socket.socket) -> "asyncio.Future[Any]":
            return asyncio.get_running_loop().create_task(self._aceitar(listener))

        async def _aceitar(self, listener: socket.socket) -> Any:
            while True:
                try:
                    return await super().accept(listener)
                except OSError as erro:
                    if not cliente_desistiu(erro) or listener.fileno() == -1:
                        raise
                    log.debug("conexão abandonada antes de ser aceita (WinError %s): aceitando a "
                              "próxima", getattr(erro, "winerror", None))  # fmt: skip


def laco() -> asyncio.AbstractEventLoop:
    """O laço de eventos do servidor (``loop`` do uvicorn, como ``módulo:função``)."""
    if sys.platform == "win32":
        return asyncio.ProactorEventLoop(_Proactor())
    return asyncio.new_event_loop()


FABRICA_DO_LACO = f"{__name__}:{laco.__name__}"


# ------------------------------------------------------------------------------------- registro
def _abandonada(mensagem: str, erro: BaseException | None) -> bool:
    if mensagem.startswith(CONEXAO_DERRUBADA):
        return (isinstance(erro, ConnectionResetError)
                and getattr(erro, "winerror", None) == WSAECONNRESET)  # fmt: skip
    # a falha que ``_Proactor`` já tratou fica também na tarefa interna do accept do asyncio
    return (
        mensagem.startswith(TAREFA_ABANDONADA)
        and TAREFA_DO_ACCEPT in mensagem
        and cliente_desistiu(erro)
    )


class _FiltroConexaoAbandonada(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        erro = record.exc_info[1] if record.exc_info else None
        if record.name != "asyncio" or not _abandonada(record.getMessage(), erro):
            return True
        record.levelno, record.levelname = logging.DEBUG, logging.getLevelName(logging.DEBUG)
        # o nível do logger já foi conferido com o de antes (ERROR): confere de novo com DEBUG
        return logging.getLogger(record.name).isEnabledFor(logging.DEBUG)


FILTRO = _FiltroConexaoAbandonada()


def instalar_filtro() -> None:
    logging.getLogger("asyncio").addFilter(FILTRO)  # não repete se já estiver
