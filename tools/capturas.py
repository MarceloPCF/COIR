"""Capturas de tela do manual e do site, com o banco de demonstração (notas fictícias do roteiro).

Cria o banco numa pasta própria e descartável na raiz do disco (``COIR-demonstracao``, com
``COIR_DADOS`` e ``COIR_CONFIG`` lá dentro, nunca a pasta do usuário; o caminho aparece nas telas,
por isso não fica na pasta temporária, que traz o nome do usuário do Windows), abre o COIR em modo
navegador num processo à parte e fotografa cada tela com o Playwright, no tema claro, em 1440x900.
Usa o Edge (ou o Chrome) instalado no Windows: não baixa navegador. As PNG vão para
docs/manual/imagens/; as telas do carrossel do site saem também em JPEG em docs/site/imagens/.

Uso: uv run python -m tools.capturas [tela ...]   (sem argumentos, todas)
"""

import os
import shutil
import socket
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
MANUAL = RAIZ / "docs" / "manual" / "imagens"
SITE = RAIZ / "docs" / "site" / "imagens"
LARGURA, ALTURA = 1440, 900
PASTA = Path(os.environ.get("SYSTEMDRIVE", "C:") + os.sep) / "COIR-demonstracao"
MARCA = ".capturas-coir"  # só apaga a pasta que ele mesmo criou
ANO = 2025  # o ano do roteiro de demonstração
AGORA = datetime(2026, 10, 16, 12, 0, 0)  # o mesmo relógio do banco de demonstração
# telas do carrossel do site: captura -> arquivo JPEG
SITE_TELAS = {"painel": "tela-painel.jpg", "apuracao": "tela-apuracao.jpg",
              "carteira": "tela-carteira.jpg", "darf": "tela-darf.jpg",
              "declaracao": "tela-declaracao.jpg", "notas": "tela-notas.jpg"}  # fmt: skip


# telas com painel de detalhe: a linha da tabela que a captura abre (texto da linha)
DETALHE = {"notas": "41004", "darf": "10/2025", "carteira": "PETR4", "opcoes": "VALEQ550"}
ESPERA = """() => ![...document.querySelectorAll('.q-spinner, .q-linear-progress--indeterminate')]
    .some(e => e.offsetParent !== null)"""


def nome(rota: str) -> str:
    """'/' -> 'painel'; '/outros-custos' -> 'outros-custos'."""
    return rota.strip("/") or "painel"


def _porta_livre() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _servidor(dados: Path, porta: int) -> None:
    """Processo do app: titular 1, ano do roteiro, relógio fixo, tema claro."""
    from coir.config import salvar_preferencia
    from coir.servicos.estado import Estado
    from coir.ui.janela import iniciar

    salvar_preferencia("tema", "claro")
    estado = Estado(dados / "coir.db")
    estado.agora = lambda: AGORA
    estado.titular, estado.ano = 1, ANO
    iniciar(estado, navegador=True, porta=porta, abrir=False)


def _esperar(porta: int, segundos: float = 60) -> None:
    fim = time.monotonic() + segundos
    while time.monotonic() < fim:
        try:
            with socket.create_connection(("127.0.0.1", porta), timeout=1):
                return
        except OSError:
            time.sleep(0.5)
    raise RuntimeError("o COIR não abriu a porta a tempo")


def _navegador(p: object) -> object:
    from playwright.sync_api import Playwright

    assert isinstance(p, Playwright)
    for canal in ("msedge", "chrome"):
        try:
            return p.chromium.launch(channel=canal)
        except Exception:  # noqa: BLE001 - tenta o próximo navegador instalado
            continue
    return p.chromium.launch()  # navegador do Playwright, se instalado (playwright install)


def fotografar(porta: int, rotas: list[str]) -> list[Path]:
    from PIL import Image
    from playwright.sync_api import Browser, sync_playwright
    from playwright.sync_api import TimeoutError as PlaywrightTimeout

    MANUAL.mkdir(parents=True, exist_ok=True)
    SITE.mkdir(parents=True, exist_ok=True)
    feitas: list[Path] = []
    with sync_playwright() as p:
        navegador = _navegador(p)
        assert isinstance(navegador, Browser)
        pagina = navegador.new_page(viewport={"width": LARGURA, "height": ALTURA},
                                    locale="pt-BR", color_scheme="light")  # fmt: skip
        for rota in rotas:
            pagina.goto(f"http://127.0.0.1:{porta}{rota}", wait_until="networkidle")
            # telas que calculam em segundo plano mostram um indicador visível até terminar
            try:
                pagina.wait_for_function(ESPERA, timeout=90_000)
            except PlaywrightTimeout:
                print(f"  aviso: {rota} ainda mostrava um indicador de carregamento")
            if nome(rota) in DETALHE:
                linhas = pagina.locator(".ag-center-cols-container .ag-row")
                linhas.filter(has_text=DETALHE[nome(rota)]).first.click()
                pagina.wait_for_load_state("networkidle")
            pagina.wait_for_timeout(800)  # gráficos (ECharts) terminam a animação
            destino = MANUAL / f"{nome(rota)}.png"
            pagina.screenshot(path=destino)
            feitas.append(destino)
            if nome(rota) in SITE_TELAS:
                Image.open(destino).convert("RGB").save(
                    SITE / SITE_TELAS[nome(rota)], quality=85, optimize=True
                )
            print(f"  {destino.relative_to(RAIZ)}")
        navegador.close()
    return feitas


def main(pedidas: list[str]) -> None:
    from coir.ui.app import INICIO, TELAS
    from tools import banco_demo, notas_ficticias

    rotas = [t.rota for t in (*TELAS, INICIO)]
    if pedidas:
        rotas = [r for r in rotas if nome(r) in pedidas]
    if PASTA.exists():
        if not (PASTA / MARCA).exists():
            raise SystemExit(
                f"{PASTA} existe e não foi criada pelas capturas: apague-a ou mude PASTA"
            )
        shutil.rmtree(PASTA)
    PASTA.mkdir()
    (PASTA / MARCA).touch()
    try:
        dados = PASTA
        banco_demo.criar(dados, roteiro=notas_ficticias.ROTEIRO_DEMONSTRACAO)
        porta = _porta_livre()
        ambiente = {**os.environ, "COIR_DADOS": str(dados),
                    "COIR_CONFIG": str(PASTA / "config.json")}  # fmt: skip
        processo = subprocess.Popen(
            [sys.executable, "-m", "tools.capturas", "--servidor", str(dados), str(porta)],
            cwd=RAIZ, env=ambiente,
        )  # fmt: skip
        try:
            _esperar(porta)
            fotografar(porta, rotas)
        finally:
            processo.terminate()
            processo.wait(timeout=30)
    finally:
        shutil.rmtree(PASTA, ignore_errors=True)


if __name__ in {"__main__", "__mp_main__"}:
    if len(sys.argv) > 1 and sys.argv[1] == "--servidor":
        _servidor(Path(sys.argv[2]), int(sys.argv[3]))
    elif __name__ == "__main__":
        main(sys.argv[1:])
