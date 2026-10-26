"""Pacote de demonstração para mostrar o COIR do zero: o programa, as notas fictícias do roteiro
e dois atalhos. "Abrir a demonstração" usa uma pasta de dados e uma configuração próprias, ao
lado dos atalhos (nunca a pasta nem a configuração de quem já usa o COIR); "Zerar a
demonstração" apaga essa pasta para começar de novo.

Uso: uv run python -m tools.pacote_demonstracao <COIR-x.y.z-windows.zip> <pasta de saída>
"""

import sys
import tempfile
import zipfile
from pathlib import Path

from tests.apoio import cpf_de_teste
from tools import notas_ficticias

ABRIR = r"""@echo off
rem Abre o COIR com a pasta de dados e a configuração desta demonstração.
set "AQUI=%~dp0"
set "COIR_DADOS=%AQUI%dados-demonstracao"
set "COIR_CONFIG=%AQUI%dados-demonstracao\config.json"
start "" "%AQUI%COIR\COIR.exe"
"""
ZERAR = r"""@echo off
rem Apaga os dados desta demonstração (só a pasta dados-demonstracao, ao lado deste arquivo).
set "AQUI=%~dp0"
if not exist "%AQUI%dados-demonstracao" (
  echo Nada a zerar.
  pause
  exit /b
)
choice /m "Apagar os dados da demonstracao e comecar do zero"
if errorlevel 2 exit /b
rmdir /s /q "%AQUI%dados-demonstracao"
echo Pronto: a demonstracao comeca do zero na proxima abertura.
pause
"""


def leia_me(titular: str, cpf: str) -> str:
    return f"""COIR - demonstração (dados fictícios)

1. Dê dois cliques em "Abrir a demonstração.bat" (feche antes o COIR, se estiver aberto).
2. Primeiros passos:
   - Pasta de dados: já vem preenchida (dados-demonstracao); clique em Continuar.
   - Titular: {titular}, CPF {cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}
     (fictício, o mesmo impresso nas notas).
   - Notas: clique em "Escolher PDFs" e selecione os 15 PDFs da pasta "notas-ficticias".
3. Em Importar > Planilhas da B3, importe "notas-ficticias\\movimentacao-2025-demonstracao.xlsx"
   (os proventos do ano).
4. Escolha o ano 2025 no topo: Painel, Apuração mensal, DARF, Carteira, Declaração anual...

Para começar de novo: "Zerar a demonstração.bat".

A demonstração não usa nem altera a pasta de dados de quem já usa o COIR neste computador.
Tudo é inventado: titular, CPF, conta e preços; os tickers e as datas são reais.
"""


def montar(programa: Path, saida: Path) -> Path:
    saida.mkdir(parents=True, exist_ok=True)
    roteiro = notas_ficticias.carregar()
    destino = saida / programa.name.replace("-windows.zip", "-demonstracao-windows.zip")
    with tempfile.TemporaryDirectory() as tmp:
        notas, planilha = notas_ficticias.gerar(roteiro, Path(tmp) / "notas-ficticias")
        with (
            zipfile.ZipFile(programa) as origem,
            zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z,
        ):
            raiz = "COIR-demonstracao"
            for item in origem.infolist():  # o programa, como no pacote normal (pasta COIR/)
                if not item.is_dir():
                    z.writestr(f"{raiz}/{item.filename}", origem.read(item))
            for arquivo in [*notas, *([planilha] if planilha else [])]:
                z.write(arquivo, f"{raiz}/notas-ficticias/{arquivo.name}")
            z.writestr(f"{raiz}/Abrir a demonstração.bat", ABRIR.replace("\n", "\r\n"))
            z.writestr(f"{raiz}/Zerar a demonstração.bat", ZERAR.replace("\n", "\r\n"))
            z.writestr(f"{raiz}/LEIA-ME.txt", leia_me(roteiro.titular, cpf_de_teste()))
    return destino


if __name__ == "__main__":
    print(montar(Path(sys.argv[1]), Path(sys.argv[2])))
