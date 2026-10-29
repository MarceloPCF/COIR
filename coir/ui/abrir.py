"""Abre uma pasta (no Explorer) ou um arquivo (no programa padrão) e traz a janela para a frente.

Quem abre é o servidor do COIR, um processo sem foco (a janela do COIR é de outro processo), e o
Windows não deixa processo sem foco pôr janela na frente: na primeira vez o Explorer abria atrás,
só na barra de tarefas. Depois de abrir, procura a janela nova pelo título (o nome da pasta ou do
arquivo) e a traz para a frente.
"""

import logging
import os
import sys
import threading
from pathlib import Path

from nicegui import ui

log = logging.getLogger(__name__)

ESPERA = 5.0  # segundos procurando a janela aberta
PASSO = 0.1
SW_RESTORE = 9
VK_MENU = 0x12  # Alt: com ele "pressionado", o Windows aceita trocar a janela da frente
KEYEVENTF_KEYUP = 0x2


def abrir(caminho: Path) -> None:
    """Abre a pasta ou o arquivo; fora do Windows (ou se não existe), só mostra o caminho."""
    if sys.platform != "win32" or not caminho.exists():
        ui.notify(str(caminho))
        return
    os.startfile(caminho)
    threading.Thread(target=_trazer_para_frente, args=(caminho.name,), daemon=True).start()


def _trazer_para_frente(titulo: str) -> None:
    if sys.platform != "win32":  # o pyright de outros sistemas não conhece o ctypes do Windows
        return
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    achadas: list[int] = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def visitar(janela: int, _: int) -> bool:
        if user32.IsWindowVisible(janela):
            tamanho = user32.GetWindowTextLengthW(janela)
            texto = ctypes.create_unicode_buffer(tamanho + 1)
            user32.GetWindowTextW(janela, texto, tamanho + 1)
            if titulo and titulo in texto.value:
                achadas.append(janela)
                return False
        return True

    pausa = threading.Event()  # nunca disparado: só a espera entre as procuras
    for _ in range(round(ESPERA / PASSO)):
        user32.EnumWindows(visitar, 0)
        if achadas:
            janela = achadas[0]
            user32.ShowWindow(janela, SW_RESTORE)
            user32.keybd_event(VK_MENU, 0, 0, 0)
            user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
            user32.SetForegroundWindow(janela)
            return
        pausa.wait(PASSO)
    log.debug("janela de %s não achada para trazer à frente", titulo)
