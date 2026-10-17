"""Abre o COIR numa janela própria (WebView2) ou, sem ele, no navegador padrão.

Só escuta em 127.0.0.1: nada da carteira sai do computador.
"""

import logging
import sys
from datetime import datetime

from nicegui import native, ui

from coir.config import configurar_registro
from coir.servicos.estado import Estado
from coir.ui.app import montar

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


def iniciar(estado: Estado, *, navegador: bool = False, porta: int | None = None,
            abrir: bool = True) -> None:  # fmt: skip
    if estado.existe:  # banco de versão anterior: migra (com backup) antes de abrir as telas
        estado.preparar(datetime.now())
    configurar_registro(estado.dados)
    nativo = not navegador and tem_webview2()
    if not navegador and not nativo:
        log.warning("WebView2 não encontrado: o COIR abre no navegador padrão")
    ui.run(
        lambda: montar(estado),
        title="COIR", host="127.0.0.1", port=porta or native.find_open_port(),
        native=nativo, window_size=(1440, 900) if nativo else None, reload=False,
        show=abrir and not nativo, language="pt-BR", dark=None, show_welcome_message=False,
        uvicorn_logging_level="warning",
    )  # fmt: skip
