"""Bloqueia dados pessoais no Git (CLAUDE.md, regras 9 e 13).

Modos:
- ``verificar_dados_pessoais.py ARQUIVO...``: hook de pre-commit e passo do CI
  (sobre ``git ls-files``).
- ``verificar_dados_pessoais.py --historico INTERVALO``: varre as linhas adicionadas e os arquivos
  incluídos em ``git log -p INTERVALO`` (ex.: ``origin/main..v3``).
- ``verificar_dados_pessoais.py --pre-push``: igual ao anterior, com o intervalo que o pre-commit
  informa no hook de pre-push.

Sai com código 1 e lista os achados se encontrar:
- CPF formatado (``000.000.000-00`` e dígitos repetidos são aceitos como exemplo);
- CPF sem máscara ou parcialmente formatado (11 dígitos) com dígito verificador válido;
- número de conta no formato de nota SINACOR (``<código>-<dígito> <conta>``) que não seja só zeros;
- termo da lista local ``%COIR_DADOS%/termos-sensiveis.txt`` (nome, e-mail, contas reais;
  um por linha);
- arquivo com extensão proibida (PDF, planilhas, bancos, backups, compactados) ou conteúdo de
  PDF/zip/planilha/SQLite; só ``.pdf``/``.xlsx`` são aceitos, e apenas em ``tests/fixtures/``;
- binário de tipo não permitido; arquivo maior que o limite (padrão 1024 KB).

Para aceitar uma linha conscientemente (dado fictício que passa no verificador), inclua nela o
comentário ``dados-pessoais: ok``. Mensagens nunca repetem o valor encontrado.
"""

import argparse
import os
import re
import subprocess
import sys
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from pathlib import Path, PurePosixPath

MARCA_ACEITE = "dados-pessoais: ok"
EXTENSOES_PROIBIDAS = frozenset(
    {
        ".pdf", ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".ofx",
        ".db", ".db-wal", ".db-shm", ".db-journal", ".sqlite", ".sqlite3", ".bak",
        ".zip", ".7z", ".rar", ".tar", ".gz",
    }
)  # fmt: skip
PASTA_FIXTURES = "tests/fixtures/"
EXTENSOES_FIXTURES = frozenset({".pdf", ".xlsx"})
BINARIOS_PERMITIDOS = frozenset({".jpg", ".jpeg", ".png", ".gif", ".ico", ".webp"})
ARQUIVOS_GERADOS = frozenset({"uv.lock"})  # hashes longos; sem varredura de conteúdo nem limite
ASSINATURAS = (
    (b"%PDF-", "PDF"),
    (b"PK\x03\x04", "zip/xlsx"),
    (b"SQLite format 3\x00", "banco SQLite"),
    (b"\xd0\xcf\x11\xe0", "planilha/documento Office antigo"),
)
LIMITE_KB_PADRAO = 1024
ARQUIVO_TERMOS = "termos-sensiveis.txt"
REF_ZERO = "0" * 40

RE_CPF_FORMATADO = re.compile(r"(?<!\d)(?<!\d\.)(\d{3})\.(\d{3})\.(\d{3})-(\d{2})(?!\d)")
# 11 dígitos, com ou sem separadores de CPF; evita pegar pedaços de valores monetários.
RE_CPF_LIVRE = re.compile(
    r"(?<!\d)(?<!\d[.,/-])(\d{3})([. ]?)(\d{3})\2(\d{3})[- ]?(\d{2})(?!\d)(?![.,/-]\d)"
)
RE_CONTA = re.compile(r"(?<![\d.,/-])\d{1,4}-\d[ \t]+(\d{6,})(?!\d)")


@dataclass(frozen=True)
class Achado:
    arquivo: str
    linha: int
    motivo: str

    def __str__(self) -> str:
        local = f"{self.arquivo}:{self.linha}" if self.linha else self.arquivo
        return f"{local}: {self.motivo}"


def cpf_valido(digitos: str) -> bool:
    """Confere os dígitos verificadores de um CPF de 11 dígitos (sem máscara)."""
    if len(digitos) != 11 or not digitos.isdigit() or len(set(digitos)) == 1:
        return False
    nums = [int(c) for c in digitos]
    for pos in (9, 10):
        soma = sum(n * peso for n, peso in zip(nums[:pos], range(pos + 1, 1, -1), strict=True))
        if (soma * 10) % 11 % 10 != nums[pos]:
            return False
    return True


def carregar_termos(caminho: Path | None = None) -> list[str]:
    """Lê a lista local de termos sensíveis (fora do repositório). Ausente => lista vazia."""
    if caminho is None:
        base = os.environ.get("COIR_DADOS", "").strip()
        caminho = (Path(base) if base else Path.home() / "COIR-dados") / ARQUIVO_TERMOS
    if not caminho.is_file():
        return []
    linhas = caminho.read_text(encoding="utf-8").splitlines()
    return [t.strip().casefold() for t in linhas if t.strip() and not t.lstrip().startswith("#")]


def verificar_texto(
    texto: str, arquivo: str = "<texto>", termos: Sequence[str] = ()
) -> list[Achado]:
    achados: list[Achado] = []
    for n, linha in enumerate(texto.splitlines(), start=1):
        if MARCA_ACEITE in linha:
            continue
        for m in RE_CPF_FORMATADO.finditer(linha):
            if len(set("".join(m.groups()))) > 1:
                achados.append(Achado(arquivo, n, "padrão de CPF formatado"))
        for m in RE_CPF_LIVRE.finditer(linha):
            digitos = m.group(1) + m.group(3) + m.group(4) + m.group(5)
            if cpf_valido(digitos) and not RE_CPF_FORMATADO.fullmatch(m.group(0)):
                achados.append(Achado(arquivo, n, "11 dígitos com DV de CPF válido"))
        for m in RE_CONTA.finditer(linha):
            if set(m.group(1)) != {"0"}:
                achados.append(Achado(arquivo, n, "número de conta de corretora"))
        minuscula = linha.casefold()
        if any(t in minuscula for t in termos):
            achados.append(Achado(arquivo, n, "termo da lista local de termos sensíveis"))
    return achados


def _ler_texto(dados: bytes) -> str | None:
    """Decodifica o conteúdo; devolve None para binários."""
    if dados.startswith((b"\xff\xfe", b"\xfe\xff")):
        return dados.decode("utf-16", errors="replace")
    if b"\x00" in dados:
        return None
    return dados.decode("utf-8", errors="replace")


def _regra_de_nome(nome: str) -> Achado | None:
    sufixo = PurePosixPath(nome).suffix.lower()
    if nome.startswith(PASTA_FIXTURES) and sufixo in EXTENSOES_FIXTURES:
        return None
    if sufixo in EXTENSOES_PROIBIDAS:
        return Achado(nome, 0, f"extensão proibida ({sufixo})")
    return None


def verificar_arquivo(
    caminho: Path, limite_kb: int = LIMITE_KB_PADRAO, termos: Sequence[str] = ()
) -> list[Achado]:
    nome = caminho.as_posix()
    if (achado := _regra_de_nome(nome)) is not None:
        return [achado]
    if not caminho.is_file() or caminho.name in ARQUIVOS_GERADOS:
        return []
    achados: list[Achado] = []
    tamanho_kb = caminho.stat().st_size / 1024
    if tamanho_kb > limite_kb:
        achados.append(Achado(nome, 0, f"arquivo grande ({tamanho_kb:.0f} KB > {limite_kb} KB)"))
    dados = caminho.read_bytes()
    sufixo = caminho.suffix.lower()
    fixture_ok = nome.startswith(PASTA_FIXTURES) and sufixo in EXTENSOES_FIXTURES
    if not fixture_ok:
        for assinatura, tipo in ASSINATURAS:
            if dados.startswith(assinatura):
                return [*achados, Achado(nome, 0, f"conteúdo de {tipo}")]
    texto = _ler_texto(dados)
    if texto is not None:
        achados.extend(verificar_texto(texto, nome, termos))
    elif sufixo not in BINARIOS_PERMITIDOS and not fixture_ok:
        tipo = sufixo or "sem extensão"
        achados.append(Achado(nome, 0, f"binário de tipo não permitido ({tipo})"))
    return achados


def verificar(
    arquivos: Iterable[Path], limite_kb: int = LIMITE_KB_PADRAO, termos: Sequence[str] = ()
) -> list[Achado]:
    return [a for arq in arquivos for a in verificar_arquivo(arq, limite_kb, termos)]


def _git(*args: str) -> str:
    return subprocess.run(
        ["git", *args], check=True, capture_output=True, text=True, encoding="utf-8"
    ).stdout


def verificar_historico(intervalo: Sequence[str], termos: Sequence[str] = ()) -> list[Achado]:
    """Varre linhas adicionadas e nomes de arquivos incluídos nos commits do intervalo."""
    achados: list[Achado] = []
    nomes = _git("log", "--format=", "--name-only", "--diff-filter=A", *intervalo).splitlines()
    for nome in nomes:
        if nome and (achado := _regra_de_nome(nome)) is not None:
            achados.append(achado)
    commit = arquivo = ""
    linha = 0
    for texto in _git("log", "-p", "--no-color", "--format=commit %h", *intervalo).splitlines():
        if texto.startswith("commit "):
            commit = texto.split()[1]
        elif texto.startswith("+++ "):
            arquivo = texto[6:] if texto.startswith("+++ b/") else texto[4:]
        elif texto.startswith("@@"):
            m = re.search(r"\+(\d+)", texto)
            linha = int(m.group(1)) - 1 if m else 0
        elif texto.startswith("+"):
            linha += 1
            if PurePosixPath(arquivo).name in ARQUIVOS_GERADOS:
                continue
            for a in verificar_texto(texto[1:], f"{commit}:{arquivo}", termos):
                achados.append(Achado(a.arquivo, linha, a.motivo))
        elif texto.startswith(" "):
            linha += 1
    return achados


def _intervalo_pre_push() -> list[str]:
    origem = os.environ.get("PRE_COMMIT_FROM_REF", "")
    destino = os.environ.get("PRE_COMMIT_TO_REF", "") or "HEAD"
    if not origem or origem == REF_ZERO:
        return [destino, "--not", "--remotes"]
    return [f"{origem}..{destino}"]


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Bloqueia dados pessoais no Git.")
    parser.add_argument("arquivos", nargs="*", type=Path)
    parser.add_argument("--limite-kb", type=int, default=LIMITE_KB_PADRAO)
    parser.add_argument(
        "--termos", type=Path, help="lista de termos sensíveis (padrão: COIR_DADOS)"
    )
    parser.add_argument("--historico", metavar="INTERVALO", help="ex.: origin/main..v3")
    parser.add_argument("--pre-push", action="store_true", help="intervalo do hook de pre-push")
    parser.add_argument("--minimo", type=int, default=0, help="falha se receber menos arquivos")
    args = parser.parse_args(argv)
    termos = carregar_termos(args.termos)

    if args.pre_push or args.historico:
        intervalo = _intervalo_pre_push() if args.pre_push else [args.historico]
        achados = verificar_historico(intervalo, termos)
        alvo = "commits de " + " ".join(intervalo)
    else:
        if len(args.arquivos) < args.minimo:
            recebidos = len(args.arquivos)
            print(f"Esperava >= {args.minimo} arquivos; recebi {recebidos}.", file=sys.stderr)
            return 2
        achados = verificar(args.arquivos, args.limite_kb, termos)
        alvo = f"{len(args.arquivos)} arquivo(s)"

    for achado in achados:
        print(achado, file=sys.stderr)
    if achados:
        total = len(achados)
        print(f"\n{total} possível(is) dado(s) pessoal(is) em {alvo}. Bloqueado.", file=sys.stderr)
        return 1
    print(f"Sem dados pessoais em {alvo} ({len(termos)} termo(s) locais).", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
