# -*- coding: utf-8 -*-
"""Monta o pacote COIR para Windows x64 com Python e Java (JRE) embutidos.

Uso:
    python montar_pacote.py <projeto> <python_orig> <py-embed.zip> <jre.zip> <saida>

    projeto      pasta com o codigo do COIR (esta pasta do repositorio)
    python_orig  instalacao (ou ambiente virtual) do Python que ja roda o COIR, com as bibliotecas
                 do requirements.txt instaladas (origem das bibliotecas)
    py-embed.zip Python embeddable da MESMA versao do python_orig (python.org, "Windows embeddable
                 package (64-bit)")
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

projeto, py_orig, embed_zip, jre_zip, saida = [os.path.abspath(a) for a in sys.argv[1:6]]

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

sp_origem = os.path.join(py_orig, "Lib", "site-packages")
sp_destino = os.path.join(py, "Lib", "site-packages")
os.makedirs(sp_destino, exist_ok=True)

# Versao do Python embutido (ex.: "315" para o 3.15), lida do proprio embeddable.
import importlib.metadata as md
import re

TAG = None
for _n in os.listdir(py):
    _m = re.fullmatch(r"python(\d)(\d+)\.zip", _n)
    if _m:
        TAG = _m.group(1) + _m.group(2)
if TAG is None:
    sys.exit("Embeddable invalido: python3xx.zip nao encontrado em " + py)

# ------------------------------------------------------------ bibliotecas
# As bibliotecas sao copiadas a partir do arquivo RECORD de cada pacote instalado, incluindo as
# dependencias (numpy.libs, pandas.libs, tzdata, etc.), em vez de uma lista fixa de nomes que
# quebra a cada atualizacao. Parte-se de requirements.txt; o resto vem das dependencias.
RAIZES = ["pandas", "numpy", "openpyxl", "tabula-py", "xlwings"]


def _norm(nome):
    return re.sub(r"[-_.]+", "-", nome).lower()


try:
    from pip._vendor.packaging.requirements import Requirement
except ImportError:  # pragma: no cover
    from packaging.requirements import Requirement

dists = {_norm(d.metadata["Name"]): d for d in md.distributions(path=[sp_origem])}
fila, fechamento = list(RAIZES), {}
while fila:
    nome = _norm(fila.pop())
    if nome in fechamento:
        continue
    if nome not in dists:
        sys.exit("FALTA no Python de origem: " + nome + " (instale o requirements.txt)")
    fechamento[nome] = dists[nome]
    for r in dists[nome].requires or []:
        req = Requirement(r)
        if req.marker is not None and not req.marker.evaluate({"extra": ""}):
            continue
        fila.append(req.name)

# pywin32: so o necessario ao xlwings (o resto sao ferramentas e o editor Pythonwin)
PYWIN32_USAR = {"win32", "win32com", "win32comext", "pythoncom.py"}


def ignorar(dist, rel):
    partes = rel.replace("\\", "/").split("/")
    if "__pycache__" in partes or rel.endswith(".pyc"):
        return True
    if dist in ("pandas", "numpy") and "tests" in partes:
        return True
    if dist == "pywin32":
        topo = partes[0]
        if topo.endswith(".dist-info") or topo in PYWIN32_USAR:
            return False
        return True
    return False


for dist, d in sorted(fechamento.items()):
    if feito("pkg-" + dist):
        continue
    n = 0
    for f in d.files or []:
        rel = str(f)
        if rel.startswith("..") or os.path.isabs(rel) or ignorar(dist, rel):
            continue  # scripts/executaveis fora da pasta site-packages
        src = os.path.join(sp_origem, rel)
        if not os.path.isfile(src):
            continue
        dst = os.path.join(sp_destino, rel)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)
        n += 1
    if n == 0:
        sys.exit("Nenhum arquivo copiado de " + dist)
    marca("pkg-" + dist)
    print("ok", dist, d.version, "(%d arquivos)" % n, flush=True)

if not feito("metadados"):
    # DLLs que ficam na raiz do python: pywin32 e xlwings
    for dll in ("pythoncom%s.dll" % TAG, "pywintypes%s.dll" % TAG):
        origem_dll = os.path.join(sp_origem, "pywin32_system32", dll)
        if not os.path.isfile(origem_dll):
            sys.exit("FALTA a DLL do pywin32: " + origem_dll)
        shutil.copy2(origem_dll, py)
    for n in os.listdir(py_orig):  # versoes antigas do xlwings gravavam as DLLs na raiz do Python
        if n.startswith("xlwings") and n.endswith(".dll"):
            shutil.copy2(os.path.join(py_orig, n), py)
    # O embeddable ignora PYTHONPATH e nao enxerga a pasta do script:
    # ".." e a pasta "sistema", onde esta o pacote "coir".
    with open(os.path.join(py, "python%s._pth" % TAG), "w", newline="\r\n") as f:
        f.write("python%s.zip\n.\n..\nLib\\site-packages\n"
                "Lib\\site-packages\\win32\nLib\\site-packages\\win32\\lib\n"
                "import site\n" % TAG)
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
