"""Configuração do COIR.

A pasta de dados reais do usuário (notas, relatórios B3, gabarito, banco) fica FORA do repositório.
O caminho vem da variável de ambiente ``COIR_DADOS``; sem ela, da pasta escolhida na interface
(gravada em ``%APPDATA%/COIR/config.json``); sem nenhuma das duas, usa ``~/COIR-dados``.
"""

import json
import logging
import os
import re
from collections.abc import Mapping
from logging.handlers import RotatingFileHandler
from pathlib import Path

from coir.erros import ConfiguracaoError

VAR_DADOS = "COIR_DADOS"
VAR_CONFIG = "COIR_CONFIG"  # outro arquivo de configuração (testes, instalação portátil)
PASTA_PADRAO = "COIR-dados"
RAIZ_REPOSITORIO = Path(__file__).resolve().parents[1]

log = logging.getLogger(__name__)


def dados_dir(*, exigir: bool = False, ambiente: Mapping[str, str] | None = None) -> Path:
    """Retorna a pasta de dados do usuário.

    Args:
        exigir: se True, levanta ``ConfiguracaoError`` quando a pasta não existe.
        ambiente: variáveis de ambiente a consultar (padrão ``os.environ``); útil em testes.
    """
    env = os.environ if ambiente is None else ambiente
    valor = env.get(VAR_DADOS, "").strip()
    escolhida = _ler_config(env).get("dados", "")
    if valor:
        pasta = Path(valor).expanduser()
        origem = VAR_DADOS
    elif isinstance(escolhida, str) and escolhida.strip():
        pasta = Path(escolhida.strip())
        origem = "configuração"
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


def arquivo_config(ambiente: Mapping[str, str] | None = None) -> Path | None:
    """``COIR_CONFIG`` ou ``%APPDATA%/COIR/config.json``; sem nenhum dos dois, None."""
    env = os.environ if ambiente is None else ambiente
    if env.get(VAR_CONFIG, "").strip():
        return Path(env[VAR_CONFIG].strip()).expanduser()
    if env.get("APPDATA", "").strip():
        return Path(env["APPDATA"].strip()) / "COIR" / "config.json"
    return None


def _ler_config(ambiente: Mapping[str, str]) -> dict[str, object]:
    arquivo = arquivo_config(ambiente)
    if arquivo is None or not arquivo.is_file():
        return {}
    try:
        dados = json.loads(arquivo.read_text(encoding="utf-8"))
    except (OSError, ValueError) as erro:
        raise ConfiguracaoError(f"configuração ilegível: {arquivo} ({erro})") from erro
    if not isinstance(dados, dict):
        raise ConfiguracaoError(f"configuração ilegível: {arquivo} (esperado um objeto JSON)")
    return dados  # type: ignore[return-value]


def salvar_dados_dir(pasta: Path, ambiente: Mapping[str, str] | None = None) -> Path:
    """Grava a pasta de dados escolhida na interface (vale nas próximas aberturas)."""
    env = os.environ if ambiente is None else ambiente
    arquivo = arquivo_config(env)
    if arquivo is None:
        raise ConfiguracaoError(f"sem lugar para a configuração: defina {VAR_CONFIG}")
    dados = {**_ler_config(env), "dados": str(pasta)}
    arquivo.parent.mkdir(parents=True, exist_ok=True)
    arquivo.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
    log.info("pasta de dados gravada na configuração: %s", arquivo)
    return arquivo


def dentro_do_repositorio(destino: Path) -> bool:
    """Saídas com dados pessoais (gabarito, relatórios) nunca vão para dentro do repositório: nem
    deste código-fonte nem de qualquer pasta versionada com git (outro clone, worktree)."""
    alvo = destino.resolve()
    if alvo == RAIZ_REPOSITORIO or RAIZ_REPOSITORIO in alvo.parents:
        return True
    return any((pasta / ".git").exists() for pasta in (alvo, *alvo.parents))


# ------------------------------------------------------------------------------- preferências
def ler_preferencia(chave: str, ambiente: Mapping[str, str] | None = None) -> str | None:
    """Preferência da interface gravada na configuração (ex.: tema)."""
    valor = _ler_config(os.environ if ambiente is None else ambiente).get(chave)
    return valor if isinstance(valor, str) else None


def salvar_preferencia(chave: str, valor: str, ambiente: Mapping[str, str] | None = None) -> None:
    env = os.environ if ambiente is None else ambiente
    arquivo = arquivo_config(env)
    if arquivo is None:
        raise ConfiguracaoError(f"sem lugar para a configuração: defina {VAR_CONFIG}")
    dados = {**_ler_config(env), chave: valor}
    arquivo.parent.mkdir(parents=True, exist_ok=True)
    arquivo.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")


# ------------------------------------------------------------------- registro de atividades
ARQUIVO_REGISTRO = Path("logs") / "coir.log"
RE_CPF = re.compile(r"\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b")


class _SemCpf(logging.Filter):
    """Última barreira: nenhum número com cara de CPF chega ao arquivo de registro."""

    def filter(self, record: logging.LogRecord) -> bool:
        record.msg = RE_CPF.sub("***", record.getMessage())
        record.args = None
        return True


def configurar_registro(dados: Path) -> Path:
    """Registro de atividades em ``<dados>/logs/coir.log`` (rotativo, 1 MB x 5)."""
    arquivo = dados / ARQUIVO_REGISTRO
    arquivo.parent.mkdir(parents=True, exist_ok=True)
    raiz = logging.getLogger()
    if any(getattr(h, "baseFilename", None) == str(arquivo) for h in raiz.handlers):
        return arquivo
    manipulador = RotatingFileHandler(arquivo, maxBytes=1_000_000, backupCount=5,
                                      encoding="utf-8")  # fmt: skip
    manipulador.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    manipulador.addFilter(_SemCpf())
    manipulador.setLevel(logging.INFO)
    raiz.addHandler(manipulador)
    logging.getLogger("alembic").setLevel(logging.WARNING)  # "Context impl" a cada conexão
    if raiz.level > logging.INFO or raiz.level == logging.NOTSET:
        raiz.setLevel(logging.INFO)
    return arquivo


def ler_registro(dados: Path, linhas: int = 300) -> list[str]:
    """As últimas linhas do registro de atividades (vazio se ainda não há registro)."""
    arquivo = dados / ARQUIVO_REGISTRO
    if not arquivo.is_file():
        return []
    return arquivo.read_text(encoding="utf-8", errors="replace").splitlines()[-linhas:]
