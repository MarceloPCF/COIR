"""Atualização das tabelas de referência pelo pacote público (F7b, "Atualizar tabelas").

Só download de dados públicos: o app consulta os releases ``referencias-AAAA.MM.DD`` do repositório
oficial no GitHub e baixa o zip; nada da carteira sai do PC. O zip só é usado se cada arquivo bater
com o SHA-256 do manifesto, o formato for o desta versão e as tabelas carregarem sem erro. A troca
é de uma vez (``referencias/atual``), e a anterior fica em ``referencias/anterior`` para voltar.
Sem rede: mensagem, nada muda (vale o pacote em uso ou o da instalação).
"""

import csv
import hashlib
import io
import json
import logging
import re
import shutil
import urllib.error
import urllib.request
import zipfile
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from coir import referencias
from coir.config import VERSAO
from coir.erros import AtualizacaoError, ConfiguracaoError

log = logging.getLogger(__name__)
REPOSITORIO = "MarceloPCF/COIR"
RELEASES = f"https://api.github.com/repos/{REPOSITORIO}/releases?per_page=100"
DOWNLOAD = f"https://github.com/{REPOSITORIO}/releases/download/{{prefixo}}"
PREFIXO = "referencias-"
MANIFESTO = "manifesto.json"
LIMITE = 50 * 1024 * 1024  # bytes do zip e do conteúdo descompactado (o pacote tem ~0,1 MB)
RE_NOME = re.compile(r"^[A-Za-z0-9_]+\.(csv|json|toml|md)$")  # arquivos de data/ (sem pasta)
Baixar = Callable[[str], bytes]


def baixar(url: str) -> bytes:
    """GET só para o GitHub (https), com limite de tamanho."""
    if not url.startswith(("https://api.github.com/", "https://github.com/")):
        raise AtualizacaoError(f"endereço fora do GitHub recusado: {url}")
    pedido = urllib.request.Request(
        url, headers={"User-Agent": f"COIR/{VERSAO}", "Accept": "application/octet-stream, */*"}
    )
    try:
        with urllib.request.urlopen(pedido, timeout=60) as resposta:
            dados: bytes = resposta.read(LIMITE + 1)
    except urllib.error.HTTPError as erro:
        motivo = (
            "limite de consultas à área de downloads atingido; tente mais tarde"
            if erro.code == 403
            else f"a área de downloads respondeu {erro.code}"
        )
        raise AtualizacaoError(f"{motivo}; nada foi alterado") from erro
    except (urllib.error.URLError, TimeoutError, OSError) as erro:
        raise AtualizacaoError(
            f"sem conexão com a área de downloads do COIR ({erro}); nada foi alterado"
        ) from erro
    if len(dados) > LIMITE:
        raise AtualizacaoError("resposta grande demais; nada foi alterado")
    return dados


@dataclass(frozen=True)
class PacoteRemoto:
    nome: str  # referencias-AAAA.MM.DD
    data: date
    url_zip: str


@dataclass(frozen=True)
class Situacao:
    em_uso: date | None
    baixado: bool  # o pacote em uso é o baixado (senão, o da instalação)
    instalacao: date | None
    anterior: date | None  # pacote para "Voltar à anterior"


@dataclass(frozen=True)
class Mudancas:
    """O que um pacote muda em relação ao que está em uso (resumo antes de aplicar)."""

    ativos_novos: int
    eventos_novos: int
    selic_nova: tuple[str, ...]
    arquivos: tuple[str, ...]


def _pastas(dados: Path) -> tuple[Path, Path, Path]:
    base = dados / referencias.PASTA_BAIXADA.parent
    return base / "atual", base / "anterior", base


def situacao(dados: Path) -> Situacao:
    atual, anterior, _ = _pastas(dados)
    em_uso = referencias.pasta_efetiva(dados)
    info = referencias.info_pacote(em_uso)
    inst = referencias.info_pacote(referencias.pasta_efetiva(None))  # a da instalação
    ant = referencias.info_pacote(anterior)
    return Situacao(em_uso=info[1] if info else None, baixado=em_uso == atual,
                    instalacao=inst[1] if inst else None,
                    anterior=ant[1] if ant and ant[0] == referencias.FORMATO else None)  # fmt: skip


def verificar(obter: Baixar = baixar) -> PacoteRemoto | None:
    """O release de referências mais novo (None se não houver nenhum)."""
    try:
        lista = json.loads(obter(RELEASES))
    except json.JSONDecodeError as erro:
        raise AtualizacaoError("resposta inesperada da área de downloads") from erro
    pacotes: list[PacoteRemoto] = []
    for r in lista if isinstance(lista, list) else []:
        if not isinstance(r, dict):
            continue
        tag = str(r.get("tag_name", ""))
        if not tag.startswith(PREFIXO) or r.get("draft") or r.get("prerelease"):
            continue
        try:
            dia = date.fromisoformat(tag.removeprefix(PREFIXO).replace(".", "-"))
        except ValueError:
            continue
        ativos = r.get("assets") if isinstance(r.get("assets"), list) else []
        urls = [str(a.get("browser_download_url", "")) for a in ativos or []
                if isinstance(a, dict) and str(a.get("name", "")).endswith(".zip")]  # fmt: skip
        # só o download do próprio release no repositório oficial
        urls = [u for u in urls if u.startswith(DOWNLOAD.format(prefixo=f"{tag}/"))]
        if urls:
            pacotes.append(PacoteRemoto(tag, dia, urls[0]))
    return max(pacotes, key=lambda p: p.data, default=None)


def mais_novo(dados: Path, remoto: PacoteRemoto | None) -> bool:
    s = situacao(dados)
    return remoto is not None and (s.em_uso is None or remoto.data > s.em_uso)


def _conferir(bruto: bytes) -> dict[str, bytes]:
    """Conteúdo do zip, conferido contra o manifesto (hash, formato, arquivo sobrando)."""
    try:
        with zipfile.ZipFile(io.BytesIO(bruto)) as z:
            itens = [i for i in z.infolist() if not i.is_dir()]
            nomes = [i.filename for i in itens]
            if len(set(nomes)) != len(nomes) or not all(RE_NOME.match(n) for n in nomes):
                raise AtualizacaoError("pacote com nome de arquivo inválido ou repetido")
            if sum(i.file_size for i in itens) > LIMITE:
                raise AtualizacaoError("pacote grande demais")
            conteudo = {i.filename: z.read(i) for i in itens}
    except (zipfile.BadZipFile, OSError, ValueError) as erro:
        raise AtualizacaoError("o pacote baixado não é um zip válido") from erro
    if MANIFESTO not in conteudo:
        raise AtualizacaoError("pacote sem manifesto")
    try:
        manifesto = json.loads(conteudo.pop(MANIFESTO))
    except ValueError as erro:
        raise AtualizacaoError("manifesto ilegível") from erro
    if not isinstance(manifesto, dict) or not isinstance(manifesto.get("arquivos"), dict):
        raise AtualizacaoError("manifesto fora do formato")
    if manifesto.get("formato") != referencias.FORMATO:
        raise AtualizacaoError(f"pacote no formato {manifesto.get('formato')}; esta versão lê o "
                               f"formato {referencias.FORMATO} (atualize o COIR)")  # fmt: skip
    esperados = {str(k): str(v) for k, v in manifesto["arquivos"].items()}
    if set(esperados) != set(conteudo):
        raise AtualizacaoError("arquivos do pacote não conferem com o manifesto")
    for nome, dados in conteudo.items():
        if hashlib.sha256(dados).hexdigest() != esperados[nome]:
            raise AtualizacaoError(f"{nome}: SHA-256 não confere; nada foi alterado")
    return conteudo


def _linhas(pasta: Path, nome: str) -> set[tuple[str, ...]]:
    caminho = pasta / nome
    if not caminho.is_file():
        return set()
    with caminho.open(encoding="utf-8", newline="") as f:
        return {tuple(r) for r in list(csv.reader(f))[1:]}


def _mudancas(atual: Path, nova: Path) -> Mudancas:
    def novas(nome: str) -> set[tuple[str, ...]]:
        return _linhas(nova, nome) - _linhas(atual, nome)

    ativos = {r[0] for r in novas("instrumentos_b3.csv")} - {
        r[0] for r in _linhas(atual, "instrumentos_b3.csv")}  # fmt: skip
    arquivos = tuple(sorted(
        p.name for p in nova.iterdir()
        if not (atual / p.name).is_file() or (atual / p.name).read_bytes() != p.read_bytes()
    ))  # fmt: skip
    return Mudancas(ativos_novos=len(ativos), eventos_novos=len(novas("eventos_b3.csv")),
                    selic_nova=tuple(sorted(r[0] for r in novas("selic_mensal.csv"))),
                    arquivos=arquivos)  # fmt: skip


def preparar(dados: Path, remoto: PacoteRemoto, obter: Baixar = baixar) -> Mudancas:
    """Baixa e confere o pacote em ``referencias/novo`` (sem usar ainda) e diz o que muda."""
    _, _, base = _pastas(dados)
    conteudo = _conferir(obter(remoto.url_zip))
    nova = base / "novo"
    try:
        if nova.exists():
            shutil.rmtree(nova)
        nova.mkdir(parents=True)
        for nome, bruto in conteudo.items():
            (nova / nome).write_bytes(bruto)
    except OSError as erro:
        raise AtualizacaoError(f"não foi possível gravar o pacote: {erro}") from erro
    info = referencias.info_pacote(nova)
    try:
        if info is None or info != (referencias.FORMATO, remoto.data):
            raise AtualizacaoError("pacote.json não confere com o release")
        # sem o cache de carregar(): a pasta "novo" é a mesma a cada pacote preparado
        referencias.carregar.__wrapped__(nova)
        referencias.carregar_fiscais.__wrapped__(nova)
        referencias.validar_tabelas_b3(nova)
    except (AtualizacaoError, ConfiguracaoError, ValueError, KeyError) as erro:
        shutil.rmtree(nova, ignore_errors=True)
        raise AtualizacaoError(f"pacote recusado: {erro}") from erro
    return _mudancas(referencias.pasta_efetiva(dados), nova)


def aplicar(dados: Path) -> None:
    """Passa o pacote preparado a ser o em uso; o anterior fica para voltar."""
    atual, anterior, base = _pastas(dados)
    nova = base / "novo"
    if referencias.info_pacote(nova) is None:
        raise AtualizacaoError("nenhum pacote preparado (verifique de novo)")
    try:
        if anterior.exists():
            shutil.rmtree(anterior)
        if atual.exists():
            atual.rename(anterior)
        try:
            nova.rename(atual)
        except OSError:
            if anterior.exists() and not atual.exists():
                anterior.rename(atual)  # desfaz: o pacote em uso continua o mesmo
            raise
    except OSError as erro:
        raise AtualizacaoError(f"não foi possível trocar as tabelas ({erro}); nada mudou") from erro
    referencias.carregar.cache_clear()
    referencias.carregar_fiscais.cache_clear()
    log.info("tabelas de referência atualizadas: %s", referencias.info_pacote(atual))


def voltar(dados: Path) -> None:
    """Troca o pacote em uso pelo anterior (e vice-versa)."""
    atual, anterior, base = _pastas(dados)
    if not anterior.is_dir():
        raise AtualizacaoError("não há pacote anterior")
    troca = base / "troca"
    try:
        if troca.exists():
            shutil.rmtree(troca)  # sobra de uma troca interrompida
        if atual.exists():
            atual.rename(troca)
        try:
            anterior.rename(atual)
        except OSError:
            if troca.exists() and not atual.exists():
                troca.rename(atual)  # desfaz
            raise
        if troca.exists():
            troca.rename(anterior)
    except OSError as erro:
        raise AtualizacaoError(f"não foi possível voltar ({erro}); nada mudou") from erro
    referencias.carregar.cache_clear()
    referencias.carregar_fiscais.cache_clear()
