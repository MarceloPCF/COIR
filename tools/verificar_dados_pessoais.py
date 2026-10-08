# -*- coding: utf-8 -*-
"""
Trava contra a publicacao de dados pessoais no repositorio do COIR.

Uso:
    python tools/verificar_dados_pessoais.py --staged   # arquivos do proximo commit (usado pelo hook)
    python tools/verificar_dados_pessoais.py --all      # todos os arquivos que o Git iria versionar
    python tools/verificar_dados_pessoais.py ARQ [ARQ]  # arquivos especificos

O que bloqueia (codigo de saida 1):
  * PDFs (notas de corretagem);
  * planilhas Excel que nao sejam o modelo vazio "modelos/COIR.xlsb";
  * qualquer arquivo dentro de Entrada/, Saida/, Resultado/ ou _dados/ (exceto .gitkeep);
  * dados/ajustes_manuais.csv e arquivos de log (log_*.txt, *.log);
  * qualquer CPF valido (confere os digitos verificadores) no conteudo dos arquivos,
    inclusive dentro de planilhas .xlsx/.xlsb.

O CPF encontrado NUNCA e impresso por inteiro: so os 3 primeiros digitos.
Somente biblioteca padrao do Python (3.8 ou superior).
"""
import os
import re
import subprocess
import sys
import zipfile
import io

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

PLANILHAS_PERMITIDAS = {"modelos/coir.xlsb"}
EXT_PLANILHA = {".xlsx", ".xlsm", ".xls", ".xlsb", ".ods"}
EXT_PDF = {".pdf"}
EXT_CONTAINER = {".xlsx", ".xlsm", ".xlsb", ".docx", ".zip"}
EXT_IMAGEM = {".jpg", ".jpeg", ".png", ".gif", ".ico"}
PASTAS_PRIVADAS = ("entrada/", "saida/", "resultado/", "_dados/")
ARQUIVOS_PRIVADOS = {"dados/ajustes_manuais.csv"}

# CPF no formato 000.000.000-00, em texto simples (ASCII/UTF-8) e em UTF-16LE (usado nas planilhas)
CPF_TEXTO = re.compile(rb"(?<!\d)\d{3}\.\d{3}\.\d{3}-\d{2}(?!\d)")
CPF_UTF16 = re.compile(rb"(?:\d\x00){3}\.\x00(?:\d\x00){3}\.\x00(?:\d\x00){3}-\x00(?:\d\x00){2}")


def cpf_valido(texto):
    """Confere os dois digitos verificadores do CPF."""
    d = [int(c) for c in texto if c.isdigit()]
    if len(d) != 11 or len(set(d)) == 1:
        return False
    for n in (9, 10):
        soma = sum(d[i] * (n + 1 - i) for i in range(n))
        if (soma * 10 % 11) % 10 != d[n]:
            return False
    return True


def cpfs_em(dados):
    """Devolve o conjunto de CPFs validos encontrados nos bytes."""
    achados = set()
    for m in CPF_TEXTO.finditer(dados):
        t = m.group().decode("ascii")
        if cpf_valido(t):
            achados.add(t)
    for m in CPF_UTF16.finditer(dados):
        t = m.group().decode("utf-16le")
        if cpf_valido(t):
            achados.add(t)
    return achados


def cpfs_no_arquivo(dados, ext):
    """Procura CPFs no conteudo; abre planilhas (que sao arquivos zip) parte por parte."""
    if ext in EXT_CONTAINER:
        try:
            with zipfile.ZipFile(io.BytesIO(dados)) as z:
                achados = set()
                for nome in z.namelist():
                    achados |= cpfs_em(z.read(nome))
                return achados
        except zipfile.BadZipFile:
            pass
    return cpfs_em(dados)


def git(args, binario=False):
    r = subprocess.run(["git"] + args, cwd=RAIZ, stdout=subprocess.PIPE,
                       stderr=subprocess.PIPE)
    if r.returncode != 0:
        raise RuntimeError(r.stderr.decode("utf-8", "replace").strip())
    return r.stdout if binario else r.stdout.decode("utf-8", "replace")


def listar(argv):
    if "--staged" in argv:
        saida = git(["diff", "--cached", "--name-only", "--diff-filter=ACMR", "-z"])
        return [p for p in saida.split("\0") if p], True
    if "--all" in argv:
        saida = git(["ls-files", "-co", "--exclude-standard", "-z"])
        return [p for p in saida.split("\0") if p], False
    arquivos = []
    for a in argv:
        if a.startswith("--"):
            continue
        caminho = os.path.abspath(a)
        arquivos.append(os.path.relpath(caminho, RAIZ).replace(os.sep, "/"))
    return arquivos, False


def ler(rel, do_indice):
    if do_indice:
        return git(["show", ":" + rel], binario=True)
    with open(os.path.join(RAIZ, rel), "rb") as f:
        return f.read()


def verificar(rel, do_indice):
    problemas = []
    baixo = rel.replace("\\", "/").lower()
    ext = os.path.splitext(baixo)[1]
    nome = os.path.basename(baixo)

    if baixo.startswith(PASTAS_PRIVADAS) and nome != ".gitkeep":
        problemas.append("arquivo dentro de pasta de dados do usuario (Entrada/Saida/Resultado)")
    if baixo in ARQUIVOS_PRIVADOS:
        problemas.append("arquivo de ajustes pessoais (contem conta, datas e quantidades reais)")
    if ext in EXT_PDF:
        problemas.append("PDF (possivel nota de corretagem)")
    if ext in EXT_PLANILHA and baixo not in PLANILHAS_PERMITIDAS:
        problemas.append("planilha que nao e o modelo vazio modelos/COIR.xlsb")
    if ext == ".log" or (nome.startswith("log_") and ext == ".txt"):
        problemas.append("log de processamento (o nome contem o CPF)")

    if ext not in EXT_IMAGEM:
        try:
            achados = cpfs_no_arquivo(ler(rel, do_indice), ext)
        except (OSError, RuntimeError) as e:
            problemas.append("nao foi possivel ler o arquivo (%s)" % e)
            achados = set()
        if achados:
            mascarados = sorted({c[:3] + ".***.***-**" for c in achados})
            problemas.append("CPF valido no conteudo (%d distinto(s): %s)"
                             % (len(achados), ", ".join(mascarados)))
    return problemas


def main(argv):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if not argv or "-h" in argv or "--help" in argv:
        print(__doc__)
        return 0
    try:
        arquivos, do_indice = listar(argv)
    except RuntimeError as e:
        print("ERRO ao consultar o Git: %s" % e)
        return 2

    bloqueados = 0
    for rel in arquivos:
        problemas = verificar(rel, do_indice)
        if problemas:
            bloqueados += 1
            print("BLOQUEADO: %s" % rel)
            for p in problemas:
                print("   - %s" % p)

    if bloqueados:
        print("\n%d arquivo(s) com risco de dados pessoais. Nada foi publicado." % bloqueados)
        print("Para tirar do commit:  git reset HEAD <arquivo>")
        print("Para o Git ignorar:    inclua o arquivo no .gitignore")
        return 1
    print("OK: %d arquivo(s) verificado(s), nenhum dado pessoal encontrado." % len(arquivos))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
