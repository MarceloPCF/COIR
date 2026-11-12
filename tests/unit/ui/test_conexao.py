"""Conexão que o navegador abandona (troca rápida de tela): o servidor continua aceitando, e o
registro de atividades não mostra como erro o que não é do COIR."""

import asyncio
import logging
import socket
import struct
import sys
import time
from collections.abc import Callable, Iterator
from pathlib import Path
from typing import Any

import pytest

from coir.ui import conexao

DERRUBADA = "Exception in callback _ProactorBasePipeTransport._call_connection_lost()"
TAREFA = (
    "Task exception was never retrieved\nfuture: <Task finished name='Task-4' coro=<{}() done, "
    "defined at windows_events.py:564> exception=OSError(22, 'rede', None, 64, None)>"
)
ACCEPT = TAREFA.format("IocpProactor.accept.<locals>.accept_coro")
OUTRA_TAREFA = TAREFA.format("Servidor.atender")
so_windows = pytest.mark.skipif(sys.platform != "win32", reason="laço Proactor do Windows")


def _erro_do_asyncio(mensagem: str, erro: OSError, winerror: int) -> None:
    """Como o laço de eventos registra a exceção de um callback (``default_exception_handler``)."""
    erro.winerror = winerror  # type: ignore[misc]  # fora do Windows o atributo não existe
    try:
        raise erro
    except OSError as levantado:
        laco = asyncio.new_event_loop()
        try:
            laco.call_exception_handler({"message": mensagem, "exception": levantado})
        finally:
            laco.close()


@pytest.fixture
def asyncio_log() -> Iterator[logging.Logger]:
    log = logging.getLogger("asyncio")
    conexao.instalar_filtro()
    yield log
    log.setLevel(logging.NOTSET)
    log.removeFilter(conexao.FILTRO)


def test_conexao_abandonada_nao_vira_erro_no_registro(
    tmp_path: Path, asyncio_log: logging.Logger
) -> None:
    """O ConnectionResetError [WinError 10054] de ``_call_connection_lost`` e a tarefa interna do
    accept abandonado descem para DEBUG e somem do registro e da auditoria. Outro erro do asyncio,
    mesmo de conexão, continua aparecendo."""
    from coir.config import configurar_registro
    from coir.servicos.manutencao import registro

    arquivo = configurar_registro(tmp_path)
    vistos: list[logging.LogRecord] = []
    coletor = logging.Handler(logging.DEBUG)
    coletor.emit = vistos.append  # type: ignore[method-assign]
    try:
        _erro_do_asyncio(DERRUBADA, ConnectionResetError(104, "derrubada"), 10054)
        _erro_do_asyncio(ACCEPT, OSError(22, "rede"), 64)
        auditoria = tmp_path / "logs" / "auditoria.jsonl"
        assert '"modulo":"asyncio"' not in auditoria.read_text(encoding="utf-8")
        _erro_do_asyncio("Accept failed on a socket", OSError(22, "rede"), 64)
        _erro_do_asyncio(DERRUBADA, ConnectionResetError(104, "outra"), 10053)
        _erro_do_asyncio(OUTRA_TAREFA, OSError(22, "rede"), 64)
        _erro_do_asyncio(ACCEPT, OSError(22, "parâmetro"), 87)
        for h in logging.getLogger().handlers:
            h.flush()
        entradas = registro(tmp_path)
        assert [(e.tipo, e.area, e.mensagem.splitlines()[0]) for e in entradas] == [
            ("Erro", "Sistema", "Task exception was never retrieved"),
            ("Erro", "Sistema", "Task exception was never retrieved"),
            ("Erro", "Sistema", DERRUBADA),
            ("Erro", "Sistema", "Accept failed on a socket"),
        ]
        assert "WinError 87" in entradas[0].detalhe and "Servidor.atender" in entradas[1].detalhe
        assert "10054" not in "\n".join(e.detalhe for e in entradas)
        assert auditoria.read_text(encoding="utf-8").count('"modulo":"asyncio"') >= 1

        asyncio_log.setLevel(logging.DEBUG)  # com --verbose elas aparecem, como DEBUG
        asyncio_log.addHandler(coletor)
        _erro_do_asyncio(DERRUBADA, ConnectionResetError(104, "derrubada"), 10054)
        _erro_do_asyncio(ACCEPT, OSError(22, "rede"), 64)
        assert [r.levelname for r in vistos if r.exc_info] == ["DEBUG", "DEBUG"]
    finally:
        asyncio_log.removeHandler(coletor)
        raiz = logging.getLogger()
        for h in [h for h in raiz.handlers if getattr(h, "baseFilename", "") == str(arquivo)]:
            raiz.removeHandler(h)
            h.close()


def test_cliente_desistiu_so_nos_erros_do_cliente() -> None:
    erro_64 = OSError(22, "rede")
    erro_64.winerror = 64  # type: ignore[misc]
    erro_87 = OSError(22, "parâmetro")
    erro_87.winerror = 87  # type: ignore[misc]
    assert conexao.cliente_desistiu(erro_64)
    assert conexao.cliente_desistiu(ConnectionResetError(104, "reset"))
    assert conexao.cliente_desistiu(ConnectionAbortedError(103, "abortada"))
    assert not conexao.cliente_desistiu(erro_87)
    assert not conexao.cliente_desistiu(PermissionError(13, "negado"))
    assert not conexao.cliente_desistiu(None)


# ------------------------------------------------------------------- o servidor continua aceitando
def _rst(porta: int) -> None:
    """Cliente que conecta e derruba a conexão (RST), como o navegador ao trocar de tela."""
    cliente = socket.create_connection(("127.0.0.1", porta))
    cliente.setsockopt(socket.SOL_SOCKET, socket.SO_LINGER, struct.pack("ii", 1, 0))
    cliente.close()


async def _servidor_responde(antes: Callable[[int], None]) -> bool:
    """Sobe um servidor, faz ``antes`` com a porta e diz se uma conexão nova ainda é atendida."""

    async def atender(_leitor: asyncio.StreamReader, escritor: asyncio.StreamWriter) -> None:
        escritor.write(b"ok")
        escritor.close()

    servidor = await asyncio.start_server(atender, "127.0.0.1", 0)
    porta = servidor.sockets[0].getsockname()[1]
    try:
        await asyncio.sleep(0.05)
        antes(porta)
        await asyncio.sleep(0.2)
        try:
            leitor, escritor = await asyncio.wait_for(
                asyncio.open_connection("127.0.0.1", porta), 2
            )
        except (OSError, TimeoutError):
            return False
        resposta = await leitor.read()
        escritor.close()
        return resposta == b"ok"
    finally:
        servidor.close()


def _rodar(fabrica: Callable[[], asyncio.AbstractEventLoop], antes: Callable[[int], None]) -> bool:
    with asyncio.Runner(loop_factory=fabrica) as executor:
        return executor.run(_servidor_responde(antes))


@so_windows
def test_conexao_derrubada_antes_do_accept_nao_derruba_o_servidor(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Duas conexões derrubadas enquanto o laço está ocupado: a da fila falha no accept com
    [WinError 64]. O laço padrão do Python 3.13 fecha o socket de escuta e recusa a próxima; o do
    COIR aceita a próxima."""

    def derrubar_com_o_laco_ocupado(porta: int) -> None:
        _rst(porta)  # a primeira já tem um accept pendente
        _rst(porta)  # a segunda fica na fila
        time.sleep(0.3)  # o laço ocupado: a da fila cai antes de ser aceita

    caplog.set_level(logging.DEBUG, logger=conexao.__name__)
    assert _rodar(conexao.laco, derrubar_com_o_laco_ocupado)
    assert "conexão abandonada antes de ser aceita (WinError 64)" in caplog.text


@so_windows
def test_accept_repete_so_nos_erros_do_cliente(monkeypatch: pytest.MonkeyPatch) -> None:
    """O accept que falha porque o cliente desistiu é refeito; outro erro segue para o asyncio, que
    fecha o socket de escuta como antes."""
    original = asyncio.IocpProactor.accept  # type: ignore[attr-defined]  # só no Windows
    falhas: list[OSError] = []

    def accept(self: Any, listener: socket.socket) -> "asyncio.Future[Any]":
        if falhas:
            futuro = asyncio.get_running_loop().create_future()
            futuro.set_exception(falhas.pop(0))
            return futuro
        return original(self, listener)

    monkeypatch.setattr(asyncio.IocpProactor, "accept", accept)  # type: ignore[attr-defined]

    falhas[:] = [OSError(22, "rede", None, 64), ConnectionResetError(10054, "reset", None, 10054),
                 ConnectionAbortedError(10053, "abortada", None, 10053)]  # fmt: skip
    assert _rodar(conexao.laco, lambda _porta: None)
    assert falhas == []

    falhas[:] = [OSError(22, "parâmetro", None, 87)]
    assert not _rodar(conexao.laco, lambda _porta: None)


def test_laco_do_servidor() -> None:
    laco = conexao.laco()
    try:
        assert isinstance(laco, asyncio.AbstractEventLoop)
        if sys.platform == "win32":
            assert isinstance(laco, asyncio.ProactorEventLoop)
            assert type(laco._proactor).__name__ == "_Proactor"  # type: ignore[attr-defined]
    finally:
        laco.close()
    assert conexao.FABRICA_DO_LACO == "coir.ui.conexao:laco"
