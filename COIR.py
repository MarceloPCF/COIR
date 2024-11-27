# -*- coding: utf-8 -*-
#
# ===== PYTHON 3 ======
# =================================================================================================
# Extração e exportação dos dados contidos nas Notas de Corretagem no padrão SINACOR
# Testado nas corretoras BTG, XP, Rico e Agora
# Para dúvidas e sugestões entrar em contato pelo e-mail: marcelo.pcf@gmail.com
# =================================================================================================

# Importações padrão de bibliotecas Python
import sys
import platform
import subprocess
from os.path import isfile, join, basename
from os import listdir, system
from datetime import datetime

# Importação das funções definidas
from Utils.funcoes import print_atencao,valida_corretora

# Importação das corretoras implementadas
import Utils.Corretoras.agora
import Utils.Corretoras.btg
import Utils.Corretoras.btg_bmf
import Utils.Corretoras.xp_rico_clear
import Utils.Corretoras.xp_rico_clear_bmf
import Utils.Corretoras.nao_validada

# =================================================================================================
# Verifica se está rodando a versão correta do Python
# =================================================================================================
if sys.version_info <= (3, 0):
    VERSAO_PYTHON = str(platform.python_version())
    MENSAGEM1 = f"Versao do interpretador python ({VERSAO_PYTHON}) inadequada.\n"
    MENSAGEM2 = "Este programa requer Python 3 (preferencialmente Python 3.9.2 ou superior).\n"

    sys.stdout.write(MENSAGEM1)
    sys.stdout.write(MENSAGEM2)
    sys.exit(1)

# =================================================================================================
# Carga de módulos opcionais
# =================================================================================================
def instalar_modulo(modulo):
    # COMANDO para instalar módulos
    comando = sys.executable + " -m" + " pip" + " install " + modulo
    print("-"*100)
    print("- O módulo", modulo,
    "não vem embutido na instalação do python e necessita de instalação específica.")

    print("- Instalando módulo opcional: ", modulo, "Aguarde....")
    subprocess.call([sys.executable, "-m", "pip", "install", modulo])
    if modulo == 'tabula-py':
        modulo = 'tabula'
    try:
        __import__(modulo)
    except ImportError as e:
        print("- Erro: Instalação de Módulo adicional", modulo, "falhou: " + str(e))
        print("- Para efetar a instalação manual, conecte-se a internet e utilize o comando abaixo")
        print(comando)
        input("- Digite <ENTER> para prosseguir")
        sys.exit(1)
    return comando

# =================================================================================================
# Lista de modulos opcionais
# PARA CADA MÓDULO NOVO, INCLUIR AS DUAS LINHAS, com a definição da váriavel modulo e o import
# =================================================================================================
MODULO = ''
try:
    MODULO='pandas==1.3.3'
    import  pandas as pd
except ImportError as e:
    print("-"*100)
    print(str(e))
    COMANDO=instalar_modulo(MODULO)

try:
    MODULO='openpyxl==3.0.9'
    import  openpyxl
except ImportError as e:
    print("-"*100)
    print(str(e))
    COMANDO=instalar_modulo(MODULO)

try:
    MODULO='xlwings==0.24.9'
    import  xlwings
except ImportError as e:
    print("-"*100)
    print(str(e))
    COMANDO=instalar_modulo(MODULO)

try:
    MODULO='tabula-py==2.3.0'
    import tabula
except ImportError as e:
    print("-"*100)
    print(str(e))
    COMANDO=instalar_modulo(MODULO)

# try:
#     MODULO="pyxlsb==1.0.10"
#     import pyxlsb
# except ImportError as e:
#     print("-"*100)
#     print(str(e))
#     COMANDO=instalar_modulo(MODULO)
# =================================================================================================
# Padrão de leitura dos arquivos PDF's contendo as Notas de Corretagem no padrão SINANCOR
# =================================================================================================
#col1str = {'dtype': str}
col1str = {'header': None}
kwargs = {
        'multiple_tables':False,
        'encoding': 'utf-8',
        'pandas_options': col1str,
        'stream':True,
        'guess':False
}

# =================================================================================================
#                  Módulo principal - SISTEMA DE CONTROLE DE OPERAÇÕES E IRPF
#    Leitura, análise, extração, formatação e conversão das Notas de Corretagem no padrão SINANCOR
# =================================================================================================
def extracao_nota_corretagem(path_origem='./Entrada/', ext='pdf'):
    resposta = ''
    arquivos = [
        join(path_origem, f) for f in listdir(path_origem)
        if isfile(join(path_origem, f)) and f.endswith(ext)
    ]
    for item in arquivos:
        filename = item
        log = []

        # Validação de notas de corretagem no padrão Sinacor
        try:
            validacao = tabula.read_pdf(filename, pandas_options={'header': None},
            guess=False, stream=True, multiple_tables=False, pages='1', silent=True,
            encoding="utf-8", area=(1.116,0.372,68.797,447.366))

            df_validacao = pd.concat(validacao,axis=1,ignore_index=True)
            df_validacao = pd.DataFrame({'NotaCorretagem': df_validacao[0].unique()})
            cell_value = df_validacao['NotaCorretagem'].iloc[0]

            if cell_value in ('NOTA DE NEGOCIAÇÃO','NOTA DE CORRETAGEM'):
                print('processando o arquivo:',basename(item))
                log.append(datetime.today().strftime('%d/%m/%Y %H:%M:%S') +
                ' - Processando o arquivo "' + basename(item) + '"\n')
            else:
                print_atencao()
                print('O arquivo','"'+basename(item).upper()+'"',
                'NÃO é uma Nota de Corretagem no Padrão Sinacor.','\n')
                #print('\033[33mO arquivo {} NÃO é uma NOTA de Corretagem
                #no Padrão Sinacor {}\033[m'.format(basename(item).upper(), '\n'))
                continue
        except ValueError:
            print_atencao()
            print('O arquivo','"'+basename(item).upper()+'"',
            'NÃO é uma Nota de Corretagem no Padrão Sinacor.','\n')
            #print('\033[33mO arquivo {} NÃO é uma NOTA de Corretagem no Padrão
            #Sinacor {}\033[m'.format(basename(item).upper(), '\n'))
            continue

        # Validação de Corretoras cadastradas e implementadas, e corretoras não validadas
        corretora = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False,
        stream=True, multiple_tables=True, pages='all', encoding="utf-8",
        area=(2.603,26.609,214.572,561.903))
        df_corretora = pd.concat(corretora,axis=0,ignore_index=True)
        corretora = tabula.read_pdf(filename, pages='1', **kwargs,
        area=(2.603,26.609,214.572,561.903))

        try:
            control,corretora,cell_value = valida_corretora(corretora)
            if control == 0:
                print('Corretora',cell_value, 'não implementada','\n')
                continue
            #elif corretora in 'XPxpRICOricoCLEARclear' and control == 1:
            if corretora in 'XPxpRICOricoCLEARclear' and control == 1:
                # Identifica o ano da nota de corretagem para a corretora XP,
                # porque houve alteração de layout das NC em 2024
                ano_pregao = tabula.read_pdf(filename, pandas_options={'dtype': str},
                guess=False, stream=True, multiple_tables=True, pages=1, encoding="utf-8",
                area=(50.947,428.028,73.259,564.134))
                ano_pregao = pd.concat(ano_pregao,axis=0,ignore_index=True)
                ano_pregao = int(ano_pregao['Data pregão'][0][6:10])
                lista_acoes = list(df_corretora[
                df_corretora['NOTA DE NEGOCIAÇÃO'].str.contains(cell_value,na=False)].index)
                if cell_value == (
                "XP INVESTIMENTOS CORRETORA DE CÂMBIO, TÍTULOS E VALORES MOBILIÁRIOS S.A."
                ):
                    cell_value = 'XP INVESTIMENTOS CORRETORA DE CÂMBIO, TÍTULOS E VALORES'
                lista_bmf = list(
                df_corretora[df_corretora['Unnamed: 0'].str.contains(cell_value,na=False)].index)
                n1 = len(lista_acoes)
                n2 = len(lista_bmf)
                if n2 >= 1:
                    page_acoes = '1'+'-'+str(n1)
                    page_bmf = (str(int(n1+1))+'-'+str(int(n1+n2)))
                    if ano_pregao > 2023:
                        Utils.Corretoras.xp_rico_clear.xp_rico_clear(
                        corretora,filename,item,log,page_acoes,page_bmf,control)
                    else:
                        Utils.Corretoras.xp_rico_clear.xp_rico_clear_old(
                        corretora,filename,item,log,page_acoes,page_bmf,control)
                else:
                    if ano_pregao > 2023:
                        Utils.Corretoras.xp_rico_clear.xp_rico_clear(
                        corretora,filename,item,log,'all')
                    else:
                        Utils.Corretoras.xp_rico_clear.xp_rico_clear_old(
                        corretora,filename,item,log,'all')
            elif corretora in 'XPxpRICOricoCLEARclear' and control == 2:
                # Identifica o ano da NC da XP, porque houve alteração de layout em 2024
                ano_pregao = tabula.read_pdf(filename, pandas_options={'dtype': str},
                guess=False, stream=True, multiple_tables=True, pages=1, encoding="utf-8",
                area=(50.947,428.028,73.259,564.134))
                ano_pregao = pd.concat(ano_pregao,axis=0,ignore_index=True)
                ano_pregao = int(ano_pregao['Data pregão'][0][6:10])
                if ano_pregao > 2023:
                    Utils.Corretoras.xp_rico_clear_bmf.xp_rico_clear_bmf(
                    corretora,filename,item,log,'all',control=2)
                else:
                    Utils.Corretoras.xp_rico_clear_bmf.xp_rico_clear_bmf_old(
                    corretora,filename,item,log,'all',control=2)
            elif corretora in 'AGORAagora' and control == 1:
                Utils.Corretoras.agora.agora(corretora,filename,item,log)
            elif corretora in 'BTGbtg' and control == 1:
                Utils.Corretoras.btg.btg(corretora,filename,item,log,'all',control=1)
            elif corretora in 'BTGbtg' and control == 2:
                Utils.Corretoras.btg_bmf.btg_bmf(corretora,filename,item,log,'all',control=2)
            elif control == 1:
                print()
                print(f'A corretora {corretora} ainda não foi validada.')
                print('Não há notas de corretagens suficientes para testá-la e implementá-la.')
                print('Todavia, será extraída com uma rotina de teste.')
                print('Dessa forma, Erros e inconsistência podem ocorrer durante o processamento.')
                print('-=' * 50)
                print()
                if resposta == '':
                    while resposta not in 'SsNn':
                        resposta = str(input('Deseja realmente continuar [S/N]: '))
                if resposta in 'Ss':
                    Utils.Corretoras.nao_validada.nao_validada(corretora, filename,item,log,'all')
                else:
                    continue
        except ValueError as e:
            print(e)
            print('ValueError - Corretora ',cell_value, 'ocorreu erro durante o processamento')
            print('das notas de corretagens','\n')
            continue

# =================================================================================================
# Mensagem de alerta para os aplicativos abertos do excel
# O sistema continuará após a confirmação do uauário
# =================================================================================================
def principal():
    print()
    print('-=' * 50)
    print(f'{"SISTEMA DE CONTROLE DE OPERAÇÕES E IRPF - COIR":^100}')
    print(f'{"Leitura, Extração e Formatação das Notas de Corretagem no padrão SINACOR":^100}')
    print('-=' * 50)
    print()
    excel_fechado = ' '
    print_atencao()
    print('Feche o Excel antes de iniciar o processamento das Notas de Corretagens.')
    print('Isso evitará erros e inconsistênica durante o processamento.\n')
    while excel_fechado not in 'SsNn':
        excel_fechado = str(input('O programa Excel está fechado [S/N]? ')).upper().strip()[0]
    if excel_fechado in 'Ss':
       #try:
       #    system("taskkill /f /im excel.exe")
       #except:
       #    print("Erro: o processo 'excel.exe' não foi encontrado.")
        result = subprocess.run(["taskkill", "/f", "/im", "excel.exe"],
        capture_output=True, text=True)
        if "não foi encontrado" in result.stderr:
            print("Cofirmado: programa Excel fechado.")
        print()
        print('-=' * 50)
        print('Iniciando o processamento das Notas de Corretagens...\n\n')
        extracao_nota_corretagem()

if __name__ == '__main__':
    principal()

print('-=' * 50)
print('Fim do processamento!','\n')
input('Pressione qualquer tecla para concluir.\n')
