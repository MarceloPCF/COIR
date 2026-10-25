"""Monta o pacote do COIR para Windows: PyInstaller em modo pasta, .zip para as GitHub Releases.

O zip leva `COIR.exe` (janela), `coir-cli.exe` (linha de comando), `_internal/` (Python e
bibliotecas), `LEIA-ME.txt`, `LICENSE`, os avisos de licença das bibliotecas e o endereço do
código-fonte da versão (AGPL). Depois de montar, o teste de fumaça roda os executáveis com o PATH
sem Python, numa pasta de dados temporária: versão, migração de um banco antigo (com backup) e a
interface respondendo. Nada da pasta de dados do usuário é lido nem gravado.

Uso: uv run python -m tools.empacotar [--sem-fumaca] [--saida dist]
"""

import hashlib
import os
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import time
import tomllib
import urllib.error
import urllib.request
import zipfile
from datetime import datetime
from importlib import metadata
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
SPEC = RAIZ / "empacotamento" / "coir.spec"
REPOSITORIO = "https://github.com/MarceloPCF/COIR"
REVISAO_ANTIGA = "0005"  # banco da v3 antes da F8: a fumaça migra dele até a última
PORTA_FUMACA = 8799


PYTHONS = ("python.exe", "python")


def versao() -> str:
    """A versão do pyproject, que tem de ser a mesma de coir.__version__."""
    from coir import __version__

    projeto = tomllib.loads((RAIZ / "pyproject.toml").read_text(encoding="utf-8"))
    declarada = projeto["project"]["version"]
    if declarada != __version__:
        raise SystemExit(f"versão divergente: pyproject {declarada} x coir {__version__}")
    return declarada


def pyinstaller(trabalho: Path) -> Path:
    subprocess.run(
        [sys.executable, "-m", "PyInstaller", str(SPEC), "--noconfirm", "--clean",
         "--distpath", str(trabalho / "dist"), "--workpath", str(trabalho / "work")],
        check=True, cwd=RAIZ,
    )  # fmt: skip
    return trabalho / "dist" / "COIR"


def avisos_de_terceiros() -> str:
    """Nome, versão e licença de cada biblioteca embutida (o texto completo de cada licença está
    nos metadados do próprio pacote, dentro de _internal)."""
    linhas = ["Bibliotecas de terceiros incluídas no COIR (nome, versão e licença declarada)", ""]
    vistos: set[str] = set()
    for d in sorted(metadata.distributions(), key=lambda d: d.metadata["Name"].lower()):
        nome = d.metadata["Name"]
        if nome.lower() in vistos or nome.lower() == "coir":
            continue
        vistos.add(nome.lower())
        licenca = d.metadata.get("License-Expression") or d.metadata.get("License") or ""
        if not licenca or len(licenca) > 80:  # texto da licença inteiro no campo: usa a classe
            classes = [c.split(" :: ")[-1] for c in d.metadata.get_all("Classifier") or []
                       if c.startswith("License ::")]  # fmt: skip
            licenca = ", ".join(classes) or "ver os metadados do pacote"
        linhas.append(f"{nome} {d.version} — {licenca}")
    return "\n".join(linhas) + "\n"


def leia_me(numero: str) -> str:
    return f"""COIR {numero} — Controle de Operações e Imposto de Renda (B3)

Para abrir: dê dois cliques em COIR.exe. Não é preciso instalar nada.
Na primeira vez, o COIR pergunta onde guardar os seus dados (notas, banco, backups e relatórios).
Escolha uma pasta FORA desta pasta do programa: assim, para atualizar, basta trocar esta pasta
pela da versão nova; os dados continuam onde estão e o banco é migrado com backup.

Linha de comando: coir-cli.exe --help

Manual, novidades e versões novas: {REPOSITORIO}/releases
Código-fonte desta versão (licença AGPL-3.0-or-later): {REPOSITORIO}/tree/v{numero}

Privacidade: o COIR roda só no seu computador. Nada da sua carteira é enviado a lugar nenhum;
as únicas consultas à internet (tabelas públicas da B3 e versão nova do programa) são feitas
quando você pede, ou ao abrir, se você ligar essa opção em Configurações.
"""


def montar_zip(pasta: Path, numero: str, saida: Path) -> Path:
    shutil.copy2(RAIZ / "LICENSE", pasta / "LICENSE.txt")
    (pasta / "LEIA-ME.txt").write_text(leia_me(numero), encoding="utf-8")
    (pasta / "AVISOS-DE-TERCEIROS.txt").write_text(avisos_de_terceiros(), encoding="utf-8")
    saida.mkdir(parents=True, exist_ok=True)
    destino = saida / f"COIR-{numero}-windows.zip"
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for arquivo in sorted(pasta.rglob("*")):
            if arquivo.is_file():
                z.write(arquivo, Path("COIR") / arquivo.relative_to(pasta))
    resumo = hashlib.sha256(destino.read_bytes()).hexdigest()
    (saida / f"{destino.name}.sha256").write_text(f"{resumo}  {destino.name}\n", encoding="ascii")
    return destino


# ------------------------------------------------------------------------- teste de fumaça
def _ambiente(dados: Path) -> dict[str, str]:
    """PATH sem Python (o pacote não pode depender do Python da máquina) e configuração
    temporária (nunca a do usuário)."""
    caminhos = [p for p in os.environ.get("PATH", "").split(os.pathsep)
                if p and not any((Path(p) / n).exists() for n in PYTHONS)]  # fmt: skip
    env = {k: v for k, v in os.environ.items() if not k.startswith(("PYTHON", "VIRTUAL_ENV", "UV"))}
    return env | {"PATH": os.pathsep.join(caminhos), "COIR_DADOS": str(dados),
                  "COIR_CONFIG": str(dados / "config.json")}  # fmt: skip


def _banco_antigo(banco: Path) -> None:
    """Banco na revisão antiga, criado pelo Alembic deste código-fonte."""
    from alembic import command

    from coir.db.sessao import _config_alembic, criar_engine  # pyright: ignore[reportPrivateUsage]

    engine = criar_engine(banco)
    with engine.begin() as conexao:
        config = _config_alembic()
        config.attributes["connection"] = conexao
        command.upgrade(config, REVISAO_ANTIGA)
    engine.dispose()


def fumaca(pasta: Path) -> None:
    from coir.db.sessao import revisao_final

    with tempfile.TemporaryDirectory(prefix="coir-fumaca-") as temporaria:
        dados = Path(temporaria)
        env = _ambiente(dados)
        cli = str(pasta / "coir-cli.exe")
        saida = subprocess.run([cli, "versao"], env=env, capture_output=True, text=True,
                               check=True, timeout=120)  # fmt: skip
        print(f"fumaça: versão -> {saida.stdout.strip()}")
        banco = dados / "coir.db"
        _banco_antigo(banco)
        subprocess.run([cli, "banco", "migrar"], env=env, check=True, timeout=300)
        with sqlite3.connect(banco) as c:
            [(revisao,)] = c.execute("SELECT version_num FROM alembic_version").fetchall()
        backups = list((dados / "backups").glob("coir-*.db"))
        if revisao != revisao_final() or len(backups) != 1:
            raise SystemExit(f"fumaça: migração {REVISAO_ANTIGA}->{revisao}, {len(backups)} backup")
        print(f"fumaça: banco {REVISAO_ANTIGA} -> {revisao}, com backup")
        ui = subprocess.Popen([cli, "ui", "--navegador", "--no-abrir", "--porta",
                               str(PORTA_FUMACA)], env=env)  # fmt: skip
        try:
            _esperar_http(f"http://127.0.0.1:{PORTA_FUMACA}/", 90)
        finally:
            ui.kill()
            ui.wait(30)
        print("fumaça: interface respondeu")


def _esperar_http(url: str, segundos: int) -> None:
    fim = time.monotonic() + segundos
    while time.monotonic() < fim:
        try:
            with urllib.request.urlopen(url, timeout=5) as resposta:
                if resposta.status == 200:
                    return
        except (urllib.error.URLError, OSError):
            time.sleep(1)
    raise SystemExit(f"fumaça: a interface não respondeu em {segundos} s")


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    saida = Path(args[args.index("--saida") + 1]) if "--saida" in args else RAIZ / "dist"
    numero = versao()
    inicio = datetime.now()
    trabalho = RAIZ / "build" / "pyinstaller"
    pasta = pyinstaller(trabalho)
    if "--sem-fumaca" not in args:
        fumaca(pasta)
    destino = montar_zip(pasta, numero, saida)
    tamanho = destino.stat().st_size / 1_048_576
    print(f"{destino} ({tamanho:.0f} MB) em {(datetime.now() - inicio).seconds} s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
