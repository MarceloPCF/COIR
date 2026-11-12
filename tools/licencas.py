"""AVISOS-DE-TERCEIROS.txt do pacote do COIR, feito a partir do que de fato foi para a pasta.

Os pacotes Python saem dos executáveis (o PYZ de COIR.exe e coir-cli.exe) e dos arquivos de
`_internal`, casados com os metadados instalados; o texto da licença de cada um vem dos próprios
metadados. O que não é pacote Python (Python, runtime da Microsoft, carregador do PyInstaller,
JavaScript e fontes do NiceGUI, WebView2) vem de `empacotamento/licencas/componentes.toml`, com o
texto oficial ao lado. Falha alto se um arquivo do pacote ficar sem licença, se um componente da
lista não estiver mais no pacote, se a lista não tiver sido conferida para as versões instaladas
ou se uma ferramenta de desenvolvimento entrar no pacote.
"""

import platform
import re
import sys
import tomllib
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from importlib import metadata
from pathlib import Path, PurePosixPath

RAIZ = Path(__file__).resolve().parents[1]
LICENCAS = RAIZ / "empacotamento" / "licencas"
EXECUTAVEIS = ("COIR.exe", "coir-cli.exe")
DLLS_DO_PYTHON = "@dlls-do-python"
# Código e dados do próprio COIR (licença AGPL, em LICENSE.txt).
DO_COIR = (
    "_internal/coir/**", "_internal/data/**", "_internal/docs/**",
    "LICENSE.txt", "LEIA-ME.txt", "AVISOS-DE-TERCEIROS.txt",
)  # fmt: skip
FREETYPE = (
    "Portions of this software are copyright © The FreeType Project (www.freetype.org). "
    "All rights reserved."
)
ARQUIVO_DE_LICENCA = re.compile(r"^(licen[cs]e|copying|notice)", re.IGNORECASE)
NOME_DO_REQUISITO = re.compile(r"^\s*([A-Za-z0-9][A-Za-z0-9._-]*)")


class AvisosIncompletos(Exception):
    """O pacote tem arquivo sem licença declarada, ou a lista de componentes está velha."""


@dataclass(frozen=True)
class Componente:
    nome: str
    versao: str
    licenca: str
    fonte: str
    arquivos: tuple[str, ...]
    caminhos: tuple[str, ...]


@dataclass(frozen=True)
class Lista:
    componentes: tuple[Componente, ...]
    proprios: tuple[str, ...]
    vigiadas: tuple[str, ...]
    sem_texto: dict[str, str]
    proibidos: frozenset[str] = frozenset()
    conferido: dict[str, str] = field(default_factory=dict[str, str])


def carregar(pasta: Path = LICENCAS) -> Lista:
    dados = tomllib.loads((pasta / "componentes.toml").read_text(encoding="utf-8"))
    componentes = tuple(
        Componente(
            c["nome"],
            c["versao"],
            c["licenca"],
            c["fonte"],
            tuple(c["arquivos"]),
            tuple(c["caminhos"]),
        )
        for c in dados["componente"]
    )
    return Lista(
        componentes, tuple(dados["proprios"]), tuple(dados["vigiadas"]),
        dict(dados["pacotes_sem_texto"]), frozenset(map(normalizar, dados["proibidos"])),
        dict(dados["conferido"]),
    )  # fmt: skip


def normalizar(nome: str) -> str:
    return re.sub(r"[-_.]+", "-", nome).lower()


def casa(caminho: str, padroes: Iterable[str]) -> bool:
    return any(PurePosixPath(caminho).full_match(p) for p in padroes)


def de_desenvolvimento(pyproject: Path = RAIZ / "pyproject.toml") -> set[str]:
    """Pacotes dos grupos de dependência (dev, videos...): nenhum pode ir no pacote."""
    projeto = tomllib.loads(pyproject.read_text(encoding="utf-8"))
    nomes: set[str] = set()
    for grupo in projeto.get("dependency-groups", {}).values():
        for requisito in grupo:
            achado = NOME_DO_REQUISITO.match(requisito) if isinstance(requisito, str) else None
            if achado:
                nomes.add(normalizar(achado.group(1)))
    return nomes


def expandir(lista: Lista, pasta_dlls: Path) -> Lista:
    """Troca "@dlls-do-python" pelo nome exato de cada arquivo da pasta DLLs do Python."""
    dlls = sorted(f"_internal/{p.name}" for p in pasta_dlls.glob("*") if p.is_file())
    componentes = tuple(
        Componente(c.nome, c.versao, c.licenca, c.fonte, c.arquivos,
                   tuple(x for p in c.caminhos for x in (dlls if p == DLLS_DO_PYTHON else [p])))
        for c in lista.componentes
    )  # fmt: skip
    return Lista(componentes, lista.proprios, lista.vigiadas, lista.sem_texto, lista.proibidos,
                 lista.conferido)  # fmt: skip


def verificar_cobertura(arquivos: Iterable[str], lista: Lista, de_pacotes: set[str]) -> list[str]:
    """Problemas de cobertura: arquivo de terceiro sem licença e componente sem arquivo."""
    problemas: list[str] = []
    usados: set[str] = set()
    for arquivo in sorted(arquivos):
        donos = [c.nome for c in lista.componentes if casa(arquivo, c.caminhos)]
        usados.update(donos)
        if casa(arquivo, lista.vigiadas):
            if not donos and not casa(arquivo, lista.proprios):
                problemas.append(f"arquivo de terceiro sem licença na lista: {arquivo}")
        elif not (donos or arquivo in de_pacotes or casa(arquivo, DO_COIR)
                  or casa(arquivo, lista.proprios)):  # fmt: skip
            problemas.append(f"arquivo sem dono conhecido: {arquivo}")
    for c in lista.componentes:
        if c.nome not in usados:
            problemas.append(f"componente da lista que não está no pacote: {c.nome}")
    return problemas


def verificar_versoes(conferido: dict[str, str], instaladas: dict[str, str]) -> list[str]:
    """A lista só vale para as versões em que foi conferida arquivo por arquivo."""
    return [f"componentes.toml conferido para {nome} {versao}; o pacote leva "
            f"{instaladas.get(nome, 'nenhum')}: revise os componentes e atualize [conferido]"
            for nome, versao in sorted(conferido.items())
            if instaladas.get(nome) != versao]  # fmt: skip


def modulos_dos_executaveis(pasta: Path) -> set[str]:
    from PyInstaller.archive.readers import CArchiveReader

    modulos: set[str] = set()
    for exe in EXECUTAVEIS:
        pyz = CArchiveReader(str(pasta / exe)).open_embedded_archive("PYZ.pyz")
        modulos.update(pyz.toc)
    return modulos


def pacotes_do_pacote(
    pasta: Path, modulos: set[str]
) -> tuple[dict[str, metadata.Distribution], set[str], list[str]]:
    """Distribuições Python dentro do pacote, arquivos de `_internal` que são delas e problemas."""
    instaladas = {normalizar(d.metadata["Name"]): d for d in metadata.distributions()
                  if d.metadata["Name"]}  # fmt: skip
    por_modulo = metadata.packages_distributions()
    achadas: dict[str, metadata.Distribution] = {}
    problemas: list[str] = []
    for topo in sorted({m.split(".")[0] for m in modulos}):
        if topo == "coir" or topo in sys.stdlib_module_names:
            continue
        nomes = por_modulo.get(topo)
        if not nomes:
            problemas.append(f"módulo sem pacote instalado conhecido: {topo}")
        for nome in nomes or []:
            if normalizar(nome) in instaladas:
                achadas[normalizar(nome)] = instaladas[normalizar(nome)]
            else:
                problemas.append(f"módulo {topo} aponta para pacote não instalado: {nome}")
    por_arquivo: dict[str, str] = {}
    for nome, d in instaladas.items():
        for f in d.files or []:
            if ".." not in f.parts:  # scripts e dados fora do site-packages nunca vão no pacote
                por_arquivo[f"_internal/{PurePosixPath(*f.parts)}"] = nome
    de_pacotes: set[str] = set()
    for arquivo in pasta.rglob("*"):
        relativo = arquivo.relative_to(pasta).as_posix()
        if arquivo.is_file() and relativo in por_arquivo:
            de_pacotes.add(relativo)
            achadas.setdefault(por_arquivo[relativo], instaladas[por_arquivo[relativo]])
    achadas.pop("coir", None)
    return achadas, de_pacotes, problemas


def licenca_declarada(d: metadata.Distribution) -> str:
    licenca = d.metadata.get("License-Expression") or d.metadata.get("License") or ""
    if licenca and len(licenca) <= 80:
        return licenca
    classes = [c.split(" :: ")[-1] for c in d.metadata.get_all("Classifier") or []
               if c.startswith("License ::")]  # fmt: skip
    return ", ".join(classes) or "ver o texto abaixo"


def ler_texto(caminho: Path) -> str:
    """Texto legal inteiro, sem trocar bytes: UTF-8 ou, nos textos antigos, Windows-1252."""
    try:
        dados = caminho.read_bytes()
    except OSError as erro:
        raise AvisosIncompletos(f"não consegui ler o texto de licença {caminho}: {erro}") from erro
    for codificacao in ("utf-8", "cp1252"):
        try:
            return dados.decode(codificacao)
        except UnicodeDecodeError:
            continue
    raise AvisosIncompletos(f"texto de licença em codificação desconhecida: {caminho}")


def textos_do_pacote(d: metadata.Distribution) -> list[tuple[str, str]]:
    """(arquivo dentro do .dist-info, texto) das licenças que a distribuição traz nos metadados."""
    arquivos = [f for f in d.files or []
                if f.parts[0].endswith(".dist-info")
                and ((len(f.parts) > 2 and f.parts[1] == "licenses")
                     or ARQUIVO_DE_LICENCA.match(f.name))]  # fmt: skip
    return [("/".join(f.parts[1:]), ler_texto(Path(str(f.locate()))))
            for f in sorted(arquivos, key=str)]  # fmt: skip


def _textos_do_componente(c: Componente, problemas: list[str]) -> list[tuple[str, str]]:
    textos: list[tuple[str, str]] = []
    for arquivo in c.arquivos:
        if arquivo == "@python":
            caminho = Path(sys.base_prefix) / "LICENSE.txt"
            if not caminho.is_file():
                problemas.append(f"licença do Python não encontrada em {caminho}")
                continue
            textos.append(("LICENSE.txt do Python", ler_texto(caminho)))
        elif arquivo == "@pyinstaller":
            try:
                textos += textos_do_pacote(metadata.distribution("pyinstaller"))
            except metadata.PackageNotFoundError:
                problemas.append("PyInstaller não instalado: sem a licença do carregador")
        else:
            textos.append((arquivo, ler_texto(LICENCAS / arquivo)))
    return textos


def _versao(nome: str) -> str:
    try:
        return metadata.version(nome)
    except metadata.PackageNotFoundError:
        return "nenhum"


def avisos(
    pasta: Path,
    versao_coir: str,
    lista: Lista | None = None,
    modulos: Callable[[Path], set[str]] = modulos_dos_executaveis,
) -> str:
    """Texto de AVISOS-DE-TERCEIROS.txt para a pasta montada pelo PyInstaller."""
    lista = expandir(lista or carregar(), Path(sys.base_prefix) / "DLLs")
    pacotes, de_pacotes, problemas = pacotes_do_pacote(pasta, modulos(pasta))
    arquivos = [a.relative_to(pasta).as_posix() for a in pasta.rglob("*") if a.is_file()]
    problemas += verificar_cobertura(arquivos, lista, de_pacotes)
    proibidos = sorted(set(pacotes) & (de_desenvolvimento() | lista.proibidos))
    if proibidos:
        problemas.append(f"ferramenta de desenvolvimento dentro do pacote: {', '.join(proibidos)}")
    instaladas = {n: _versao(n) for n in lista.conferido if n != "python"}
    instaladas["python"] = ".".join(platform.python_version_tuple()[:2])
    problemas += verificar_versoes(lista.conferido, instaladas)
    blocos: list[tuple[str, list[tuple[str, str]]]] = []
    resumo_py: list[str] = []
    for nome in sorted(pacotes):
        d = pacotes[nome]
        titulo = f"{d.metadata['Name']} {d.version}"
        textos = textos_do_pacote(d)
        if not textos and nome in lista.sem_texto:
            textos = [(lista.sem_texto[nome], ler_texto(LICENCAS / lista.sem_texto[nome]))]
        if not textos:
            problemas.append(f"pacote sem texto de licença: {titulo}")
        resumo_py.append(f"  {titulo} — {licenca_declarada(d)}")
        blocos.append((titulo, textos))
    resumo_outros: list[str] = []
    for c in lista.componentes:
        versao = _versao("pyinstaller") if "@pyinstaller" in c.arquivos else c.versao
        titulo = f"{c.nome} {versao}"
        resumo_outros.append(f"  {titulo} — {c.licenca}\n      {c.fonte}")
        blocos.append((titulo, _textos_do_componente(c, problemas)))
    if problemas:
        raise AvisosIncompletos("AVISOS-DE-TERCEIROS incompleto:\n  " + "\n  ".join(problemas))
    return _formatar(versao_coir, resumo_py, resumo_outros, blocos)


def _formatar(
    versao_coir: str,
    resumo_py: list[str],
    resumo_outros: list[str],
    blocos: list[tuple[str, list[tuple[str, str]]]],
) -> str:
    partes = [
        f"AVISOS DE TERCEIROS — COIR {versao_coir}",
        "",
        "O COIR é software livre, distribuído sob a GNU Affero General Public License, versão 3 ou",
        "posterior (LICENSE.txt). Este pacote também traz programas, bibliotecas e fontes de",
        "terceiros, cada um sob a sua própria licença. A parte 1 lista todos; a parte 2 reproduz o",
        "texto integral de cada licença, com os avisos de copyright.",
        "",
        f"{FREETYPE} (FreeType, dentro do PDFium e do Pillow.)",
        "",
        "PARTE 1 — COMPONENTES",
        "",
        "Pacotes Python:",
        *resumo_py,
        "",
        "Outros componentes:",
        *resumo_outros,
        "",
        "PARTE 2 — TEXTOS DAS LICENÇAS",
    ]
    for titulo, textos in blocos:
        partes += ["", "=" * 78, titulo, "=" * 78]
        for arquivo, texto in textos:
            partes += ["", f"--- {arquivo} ---", "", texto.strip()]
    return "\n".join(partes) + "\n"
