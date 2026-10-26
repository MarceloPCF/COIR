"""Janela do Windows para escolher uma pasta (o "Procurar..." do Explorer).

Na janela própria do COIR (WebView2) usa o diálogo do pywebview; no modo navegador, o diálogo de
pastas do Windows por um PowerShell local (o servidor roda no mesmo computador). Sem nenhum dos
dois, devolve None e a tela continua aceitando o caminho digitado.
"""

import logging
import subprocess
import sys
from pathlib import Path

from nicegui import app, run

log = logging.getLogger(__name__)

_POWERSHELL = (
    "Add-Type -AssemblyName System.Windows.Forms;"
    "$d = New-Object System.Windows.Forms.FolderBrowserDialog;"
    "$d.Description = 'Escolha a pasta';"
    "$d.UseDescriptionForTitle = $true;"
    "$d.SelectedPath = $env:COIR_PASTA_INICIAL;"
    "$janela = New-Object System.Windows.Forms.Form -Property @{TopMost = $true};"
    "if ($d.ShowDialog($janela) -eq 'OK') { [Console]::Out.Write($d.SelectedPath) }"
)


async def escolher_pasta(inicial: Path | None = None) -> Path | None:
    """A pasta escolhida, ou None (cancelado ou sem diálogo disponível)."""
    janela = getattr(app.native, "main_window", None)
    if janela is not None:
        import webview

        tipo = webview.FileDialog.FOLDER  # pywebview 6
        escolhidas = await janela.create_file_dialog(tipo, directory=str(inicial or ""))
        return Path(escolhidas[0]) if escolhidas else None
    if sys.platform != "win32":
        return None
    return await run.io_bound(_powershell, inicial)


def _powershell(inicial: Path | None) -> Path | None:
    import os

    ambiente = {**os.environ, "COIR_PASTA_INICIAL": str(inicial or "")}
    try:
        saida = subprocess.run(
            ["powershell", "-NoProfile", "-STA", "-Command", _POWERSHELL], env=ambiente,
            capture_output=True, text=True, timeout=600,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )  # fmt: skip
    except (OSError, subprocess.TimeoutExpired) as erro:
        log.warning("diálogo de pasta indisponível: %s", erro)
        return None
    escolhida = saida.stdout.strip()
    return Path(escolhida) if escolhida else None
