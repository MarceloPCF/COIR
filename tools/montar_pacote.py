# -*- coding: utf-8 -*-
"""Monta o pacote COIR para Windows x64 com Python e Java (JRE) embutidos.

Uso:
    python montar_pacote.py <projeto> <python39> <py-embed.zip> <jre.zip> <saida>

    projeto      pasta com o codigo do COIR (esta pasta do repositorio)
    python39     instalacao do Python 3.9.x que ja roda o COIR (origem das bibliotecas)
    py-embed.zip Python embeddable da mesma versao (python.org, "Windows embeddable package (64-bit)")
    jre.zip      Eclipse Temurin 17 JRE para Windows x64, em .zip (adoptium.net)
    saida        pasta onde serao criados COIR-v<versao>/ e o .zip

Estrutura gerada (o usuario ve so o que usa):
    COIR.bat  LEIA-ME.txt  LICENSE
    Entrada/  Saida/  Resultado/  dados/  modelos/
    sistema/  (python/, jre/, COIR.py, coir/)

O programa continua usando caminhos relativos (./Entrada, ./dados, ./modelos ...);
o COIR.bat roda a partir da pasta principal, entao nenhum codigo precisa mudar.

A montagem e retomavel: cada etapa concluida deixa uma marca em <saida>/.etapas.
Apague a pasta <saida> para refazer tudo do zero.
"""
import os
import shutil
import sys
import zipfile

projeto, py39, embed_zip, jre_zip, saida = [os.path.abspath(a) for a in sys.argv[1:6]]

with open(os.path.join(projeto, "VERSION"), encoding="utf-8") as f:
    VERSAO = f.read().strip()
NOME = "COIR-v" + VERSAO
raiz = os.path.join(saida, NOME)
sistema = os.path.join(raiz, "sistema")
py = os.path.join(sistema, "python")
os.makedirs(sistema, exist_ok=True)
FEITO = os.path.join(saida, ".etapas")
os.makedirs(FEITO, exist_ok=True)


def feito(nome):
    return os.path.exists(os.path.join(FEITO, nome))


def marca(nome):
    open(os.path.join(FEITO, nome), "w").close()


def sem_cache(pasta, nomes):
    return {n for n in nomes if n == "__pycache__" or n.endswith(".pyc")}


def limpar(destino):
    if os.path.isdir(destino):
        shutil.rmtree(destino)
    elif os.path.exists(destino):
        os.remove(destino)


# ------------------------------------------------------------ 1. programa
if not feito("programa"):
    shutil.copy2(os.path.join(projeto, "COIR.py"), sistema)
    limpar(os.path.join(sistema, "coir"))
    shutil.copytree(os.path.join(projeto, "coir"), os.path.join(sistema, "coir"), ignore=sem_cache)
    for pasta in ("dados", "modelos"):
        limpar(os.path.join(raiz, pasta))
        shutil.copytree(os.path.join(projeto, pasta), os.path.join(raiz, pasta), ignore=sem_cache)
    shutil.copy2(os.path.join(projeto, "LICENSE"), raiz)
    for pasta in ("Entrada", "Saida", "Resultado"):
        os.makedirs(os.path.join(raiz, pasta), exist_ok=True)
    marca("programa")

# ------------------------------------------------------------ 2. python
if not feito("embed"):
    with zipfile.ZipFile(embed_zip) as z:
        z.extractall(py)
    marca("embed")

sp_origem = os.path.join(py39, "Lib", "site-packages")
sp_destino = os.path.join(py, "Lib", "site-packages")
os.makedirs(sp_destino, exist_ok=True)

PACOTES = [
    "numpy", "pandas", "pytz", "dateutil", "six.py",
    "openpyxl", "et_xmlfile", "jdcal.py",
    "tabula", "distro.py",
    "xlwings",
    "win32", "win32com", "win32comext", "pythoncom.py",
    "pkg_resources",
]


def ignorar_libs(pasta, nomes):
    out = set()
    for n in nomes:
        if n == "__pycache__" or n.endswith(".pyc"):
            out.add(n)
        elif n == "tests" and ("pandas" in pasta or "numpy" in pasta):
            out.add(n)
        elif n in ("Demos", "test", "scripts", "include") and "win32" in pasta:
            out.add(n)
    return out


for nome in PACOTES:
    if feito("pkg-" + nome):
        continue
    src = os.path.join(sp_origem, nome)
    dst = os.path.join(sp_destino, nome)
    limpar(dst)
    if os.path.isdir(src):
        shutil.copytree(src, dst, ignore=ignorar_libs)
    elif os.path.isfile(src):
        shutil.copy2(src, dst)
    else:
        sys.exit("FALTA no Python de origem: " + nome)
    marca("pkg-" + nome)
    print("ok", nome, flush=True)

if not feito("metadados"):
    # dist-info / egg-info: nome e versao das bibliotecas (util para diagnostico)
    LISTA = ("numpy", "pandas", "pytz", "python-dateutil", "openpyxl", "et-xmlfile",
             "jdcal", "tabula-py", "pywin32", "distro", "six")
    for n in os.listdir(sp_origem):
        src = os.path.join(sp_origem, n)
        if n.endswith(".dist-info") and n.split("-")[0].lower().replace("_", "-") in LISTA:
            shutil.copytree(src, os.path.join(sp_destino, n), dirs_exist_ok=True)
        elif n.startswith("xlwings") and n.endswith(".egg-info"):
            if os.path.isdir(src):
                shutil.copytree(src, os.path.join(sp_destino, n), dirs_exist_ok=True)
            else:
                shutil.copy2(src, os.path.join(sp_destino, n))
    # DLLs que ficam na raiz do python: pywin32 e xlwings
    for dll in ("pythoncom39.dll", "pywintypes39.dll"):
        shutil.copy2(os.path.join(sp_origem, "pywin32_system32", dll), py)
    for n in os.listdir(py39):
        if n.startswith("xlwings") and n.endswith(".dll"):
            shutil.copy2(os.path.join(py39, n), py)
    # O embeddable ignora PYTHONPATH e nao enxerga a pasta do script:
    # ".." e a pasta "sistema", onde esta o pacote "coir".
    with open(os.path.join(py, "python39._pth"), "w", newline="\r\n") as f:
        f.write("python39.zip\n.\n..\nLib\\site-packages\n"
                "Lib\\site-packages\\win32\nLib\\site-packages\\win32\\lib\n"
                "import site\n")
    marca("metadados")

# ------------------------------------------------------------ 3. java
if not feito("jre"):
    destino_jre = os.path.join(sistema, "jre")
    limpar(destino_jre)
    with zipfile.ZipFile(jre_zip) as z:
        topo = z.namelist()[0].split("/")[0]
        z.extractall(saida)
    shutil.move(os.path.join(saida, topo), destino_jre)
    marca("jre")

# ------------------------------------------------------------ 4. COIR.bat e LEIA-ME
BAT = r"""@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title COIR - Controle de Operacoes e Imposto de Renda

if not exist "sistema\python\python.exe" (
    echo.
    echo Nao foi possivel encontrar o Python do COIR.
    echo Extraia TODO o conteudo do arquivo zip para uma pasta
    echo e execute o COIR.bat que esta nessa pasta.
    echo.
    pause
    exit /b 1
)

rem Janela inteira em azul fosco com texto branco (Windows 10 ou 11).
rem Em outros sistemas a janela fica com as cores padrao.
set "ESC="
ver | find "10.0." >nul && (
    for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do set "ESC=%%b"
)
if defined ESC <nul set /p "=%ESC%[48;2;38;70;110m%ESC%[97m%ESC%[2J%ESC%[H"

set "JAVA_HOME=%~dp0sistema\jre"
set "PATH=%~dp0sistema\jre\bin;%~dp0sistema\python;%PATH%"
set "PYTHONIOENCODING=utf-8"

"sistema\python\python.exe" "sistema\COIR.py"
set "RC=%errorlevel%"
if not "%RC%"=="0" (
    echo.
    pause
)
if defined ESC <nul set /p "=%ESC%[0m"
exit /b %RC%
"""
with open(os.path.join(raiz, "COIR.bat"), "w", newline="\r\n") as f:
    f.write(BAT)

LEIA = """COIR - Controle de Operacoes e Imposto de Renda  (versao %s)
=====================================================================

Como usar
---------
1. Extraia TODO o conteudo do zip para uma pasta (nao execute de dentro do zip).
2. Baixe suas notas de corretagem em PDF (padrao SINACOR) no portal da corretora
   e copie para a pasta  Entrada.
3. De dois cliques em  COIR.bat.
4. As notas processadas sao movidas para a pasta  Saida.
   O resultado fica em  Resultado\\CPF.xlsb  (CPF = numero do CPF do investidor).

Requisitos
----------
Windows de 64 bits e Microsoft Excel instalado.
O Python e o Java ja vem dentro da pasta  sistema : nao e preciso instalar nada.

Pastas
------
Entrada    notas de corretagem a processar
Saida      notas ja processadas
Resultado  planilhas com o resultado (uma por CPF)
dados      tabelas auxiliares (acoes, opcoes, corretoras, ajustes manuais)
modelos    planilha-modelo vazia
sistema    programa, Python e Java (nao precisa mexer)

Se o Windows alertar sobre o python.exe (SmartScreen ou antivirus), e um aviso
comum para programas baixados da internet; ele faz parte deste pacote.

Corretoras testadas: XP, Clear, Rico, Necton e BTG. Se voce opera em outra,
entre em contato para que ela seja adicionada.

Contato:  marcelo.pcf@gmail.com
Projeto:  https://github.com/MarceloPCF/COIR
Pagina:   https://marcelopcf.github.io/COIR/
Licenca:  MIT (arquivo LICENSE)
""" % VERSAO
with open(os.path.join(raiz, "LEIA-ME.txt"), "w", encoding="utf-8-sig", newline="\r\n") as f:
    f.write(LEIA)

# ------------------------------------------------------------ 5. zip
# Seguranca: o zip NUNCA leva o conteudo das pastas de trabalho (Entrada, Saida, Resultado),
# nem caches do Python, mesmo que a pasta tenha sido usada para testes com notas reais.
import re

PASTAS_TRABALHO = {"Entrada", "Saida", "Resultado"}
CPF = re.compile(r"\d{3}\.?\d{3}\.?\d{3}-?\d{2}")


def incluir(rel):
    partes = rel.replace("\\", "/").split("/")  # [COIR-vX, ...]
    if "__pycache__" in partes or rel.endswith(".pyc"):
        return False
    if len(partes) > 2 and partes[1] in PASTAS_TRABALHO:
        return False  # so as pastas vazias entram
    return True


zip_path = os.path.join(saida, NOME + ".zip")
limpar(zip_path)
with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
    for pasta, dirs, arquivos in os.walk(raiz):
        rel = os.path.relpath(pasta, saida)
        if not incluir(rel + "/x") and "__pycache__" in rel.split(os.sep):
            dirs[:] = []
            continue
        z.write(pasta, rel + "/")  # grava tambem as pastas vazias (Entrada, Saida, Resultado)
        if os.path.basename(pasta) in PASTAS_TRABALHO and os.path.dirname(pasta) == raiz:
            dirs[:] = []  # nao desce nas pastas de trabalho
            continue
        for a in arquivos:
            r = os.path.join(rel, a)
            if incluir(r):
                z.write(os.path.join(pasta, a), r)

# Conferencia final: aborta (e apaga o zip) se algo pessoal foi parar no pacote.
problemas = []
with zipfile.ZipFile(zip_path) as z:
    for i in z.infolist():
        n = i.filename
        partes = n.split("/")
        base = partes[-1].lower()
        if len(partes) > 2 and partes[1] in PASTAS_TRABALHO and not i.is_dir():
            problemas.append("arquivo em pasta de trabalho: " + n)
        if "site-packages" not in n and base.endswith((".pdf", ".xlsx", ".xlsm", ".xls", ".ods")):
            problemas.append("documento: " + n)
        if base.endswith(".xlsb") and n != NOME + "/modelos/COIR.xlsb":
            problemas.append("planilha: " + n)
        if base.startswith("log_") or base == "ajustes_manuais.csv":
            problemas.append("arquivo pessoal: " + n)
        if "site-packages" not in n and "/jre/" not in n and CPF.search(n):
            problemas.append("CPF no nome: " + n)
if problemas:
    os.remove(zip_path)
    sys.exit("ZIP ABORTADO, dados pessoais ou lixo no pacote:\n  " + "\n  ".join(problemas[:20]))
print("zip:", zip_path, round(os.path.getsize(zip_path) / 1e6, 1), "MB (conferido: sem dados pessoais)")
