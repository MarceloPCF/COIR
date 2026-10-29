"""Abre o COIR numa janela própria (WebView2) ou, sem ele, no navegador padrão.

Só escuta em 127.0.0.1: nada da carteira sai do computador.
"""

import logging
import os
import sys
from datetime import datetime
from pathlib import Path

from nicegui import app, native, ui

from coir.config import VERSAO, configurar_registro
from coir.erros import CoirError
from coir.servicos.estado import Estado
from coir.ui import tema
from coir.ui.app import montar, montar_erro

log = logging.getLogger(__name__)
WEBVIEW2 = (
    r"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients"
    r"\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
)


def tem_webview2() -> bool:
    """WebView2 instalado (máquina ou usuário). Fora do Windows, sempre o navegador."""
    if sys.platform != "win32":
        return False
    import winreg

    chaves = ((winreg.HKEY_LOCAL_MACHINE, WEBVIEW2),
              (winreg.HKEY_CURRENT_USER, WEBVIEW2.replace("WOW6432Node\\", "")))  # fmt: skip
    for raiz, caminho in chaves:
        try:
            with winreg.OpenKey(raiz, caminho) as chave:
                versao, _ = winreg.QueryValueEx(chave, "pv")
                if versao and versao != "0.0.0.0":
                    return True
        except OSError:
            continue
    return False


def desbloquear_dlls(pasta: Path) -> int:
    """Tira a marca "baixado da internet" (Zone.Identifier) das DLLs do programa. O Windows a
    copia do .zip para cada arquivo extraído, e com ela o .NET recusa carregar Python.Runtime.dll:
    a janela (pywebview) morre sem aviso e só o servidor fica rodando. Devolve quantas tirou."""
    tiradas = 0
    for dll in pasta.rglob("*.dll"):
        try:
            os.remove(f"{dll}:Zone.Identifier")
            tiradas += 1
        except OSError:  # sem a marca, ou pasta sem permissão de escrita
            continue
    return tiradas


def iniciar(estado: Estado, *, navegador: bool = False, porta: int | None = None,
            abrir: bool = True) -> None:  # fmt: skip
    configurar_registro(estado.dados)
    log.info("COIR %s aberto (pasta de dados: %s)", VERSAO, estado.dados)
    if sys.platform == "win32" and getattr(sys, "frozen", False):
        if tiradas := desbloquear_dlls(Path(sys.executable).parent):
            log.info("marca de arquivo baixado tirada de %d DLL(s) do programa", tiradas)
    erro: str | None = None
    if estado.existe:  # banco de versão anterior: migra (com backup) antes de abrir as telas
        try:
            estado.preparar(datetime.now())
        except CoirError as falha:  # a janela abre e explica (banco mais novo, migração falhou)
            log.exception("o banco não abriu")
            erro = str(falha)
    nativo = not navegador and tem_webview2()
    if not navegador and not nativo:
        log.warning("WebView2 não encontrado: o COIR abre no navegador padrão")
    # o símbolo do COIR na barra de título e na barra de tarefas; sem isso a janela usa o ícone
    # do executável, que rodando pelo Python é o do Python
    app.native.start_args["icon"] = str(tema.ICONE_JANELA)
    ui.run(
        (lambda: montar_erro(estado, erro)) if erro else (lambda: montar(estado)),
        title="COIR", favicon=tema.ICONE, host="127.0.0.1", port=porta or native.find_open_port(),
        native=nativo, window_size=(1440, 900) if nativo else None, reload=False,
        show=abrir and not nativo, language="pt-BR", dark=None, show_welcome_message=False,
        uvicorn_logging_level="warning",
        # cálculo pesado (apuração do histórico inteiro) pode segurar o servidor alguns segundos:
        # a janela espera reconectar em vez de perder a página ("conexão perdida")
        reconnect_timeout=30,
    )  # fmt: skip
