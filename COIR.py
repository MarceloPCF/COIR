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
from os.path import isfile, join, basename, exists
from os import listdir
from datetime import datetime
import pandas as pd

# Importação das funções definidas
from coir.funcoes import print_atencao, valida_corretora
import coir.funcoes
from coir import __version__

# Importação das corretoras implementadas
import coir.corretoras.agora
import coir.corretoras.btg
import coir.corretoras.btg_bmf
import coir.corretoras.xp_rico_clear
import coir.corretoras.xp_rico_clear_bmf
import coir.corretoras.nao_validada

# =================================================================================================
# Verifica se está rodando a versão correta do Python
# =================================================================================================
VERSAO_MINIMA = (3, 9, 2)
if sys.version_info < VERSAO_MINIMA:
    VERSAO_PYTHON = str(platform.python_version())
    MENSAGEM1 = f"Versao do interpretador python ({VERSAO_PYTHON}) inadequada.\n"
    MENSAGEM2 = "Este programa requer Python 3.9.2 ou superior.\n"

    sys.stdout.write(MENSAGEM1)
    sys.stdout.write(MENSAGEM2)
    sys.exit(1)

# =================================================================================================
# Carga de módulos opcionais
# =================================================================================================
def instalar_modulo(modulo):
    # COMANDO para instalar módulos
    comando = sys.executable + " -m" + " pip" + " install " + modulo
    print("-" * 100)
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


# =================================================================================================
# Lista de módulos opcionais
# PARA CADA MÓDULO NOVO, INCLUIR AS DUAS LINHAS, com a definição da váriavel modulo e o import
# =================================================================================================
MODULO = ''
#try:
#    MODULO='pandas==1.3.3'
#    import  pandas as pd
#except ImportError as e:
#    print("-"*100)
#    print(str(e))
#    instalar_modulo(MODULO)    
try:
    MODULO = 'openpyxl==3.0.9'
    import openpyxl
except ImportError as e:
    print("-" * 100)
    print(str(e))
    instalar_modulo(MODULO)

try:
    MODULO = 'xlwings==0.24.9'
    import xlwings
except ImportError as e:
    print("-" * 100)
    print(str(e))
    instalar_modulo(MODULO)

try:
    MODULO = 'tabula-py==2.3.0'
    import tabula
except ImportError as e:
    print("-" * 100)
    print(str(e))
    instalar_modulo(MODULO)

# numpy==1.21.5

# try:
#     MODULO="pyxlsb==1.0.10"
#     import pyxlsb
# except ImportError as e:
#     print("-"*100)
#     print(str(e))
#     instalar_modulo(MODULO)
# =================================================================================================
# Padrão de leitura dos arquivos PDF's contendo as Notas de Corretagem no padrão SINANCOR
# =================================================================================================
col1str = {'header': None}
kwargs = {
    'multiple_tables': False,
    'encoding': 'utf-8',
    'pandas_options': col1str,
    'stream': True,
    'guess': False
}

# ---------------------------------------------------------------------------
# Funções auxiliares
# ---------------------------------------------------------------------------

def _valida_nota_sinacor(filename):
    try:
        validacao = tabula.read_pdf(
            filename, pandas_options={'header': None}, guess=False, stream=True,
            multiple_tables=False, pages='1', silent=True, encoding="utf-8",
            area=(1.116, 0.372, 68.797, 447.366)
        )
        df_validacao = pd.concat(validacao, axis=1, ignore_index=True)
        df_validacao = pd.DataFrame({'NotaCorretagem': df_validacao[0].unique()})
        return df_validacao['NotaCorretagem'].iloc[0] in ('NOTA DE NEGOCIAÇÃO', 'NOTA DE CORRETAGEM')
    except ValueError:
        return False


def _identifica_ano_pregao(filename):
    tabelas = tabula.read_pdf(
        filename, pandas_options={'dtype': str}, guess=False, stream=True,
        multiple_tables=True, pages=1, encoding="utf-8",
        area=(50.947, 428.028, 73.259, 564.134)
    )
    df_ano = pd.concat(tabelas, axis=0, ignore_index=True)
    return int(df_ano['Data pregão'][0][6:10])


def _identifica_grupo(corretora):
    nome = corretora.upper()
    if nome in ('XP', 'RICO', 'CLEAR'):
        return 'XP'
    if nome == 'AGORA':
        return 'AGORA'
    if nome == 'BTG':
        return 'BTG'
    return None


def _append_resultado(normal_df, daytrade_df, normal_dfs, daytrade_dfs):
    if normal_df is not None and not normal_df.empty:
        normal_dfs.append(normal_df)
    if daytrade_df is not None and not daytrade_df.empty:
        daytrade_dfs.append(daytrade_df)


def aplicar_ajustes_manuais(df, caminho='./dados/ajustes_manuais.csv'):
    """Aplica correções manuais (mudança de ticker, eventos, transferência de custódia).
    Cada linha do CSV localiza uma operação por corretora, conta, data, C/V, papel, quantidade
    e total, e substitui apenas os campos preenchidos em nova_*."""
    if df.empty or not exists(caminho):
        return df
    aj = pd.read_csv(caminho, dtype=str, keep_default_na=False)
    datas = pd.to_datetime(df['Data']).dt.strftime('%Y-%m-%d')
    for _, r in aj.iterrows():
        mask = ((df['Corretora'] == r['corretora']) & (df['Conta'].astype(str) == r['conta'])
                & (datas == r['data']) & (df['C/V'] == r['cv']) & (df['Papel'] == r['papel'])
                & (df['Quantidade'].astype(float) == float(r['quantidade']))
                & ((df['Total'].astype(float) - float(r['total'])).abs() < 0.01))
        if not mask.any():
            continue
        if r['nova_corretora']:
            df.loc[mask, 'Corretora'] = r['nova_corretora']
        if r['nova_conta']:
            df.loc[mask, 'Conta'] = int(r['nova_conta'])
        if r['novo_papel']:
            df.loc[mask, 'Papel'] = r['novo_papel']
        if r.get('novo_exercicio'):
            col_ex = 'Exercicio' if 'Exercicio' in df.columns else 'Exercício'
            if col_ex in df.columns:
                df.loc[mask, col_ex] = r['novo_exercicio']
        if r['nova_quantidade']:
            q_old = float(r['quantidade']); q_new = float(r['nova_quantidade'])
            df.loc[mask, 'PM'] = df.loc[mask, 'PM'] * q_old / q_new
            df.loc[mask, 'Quantidade'] = q_new
        if r['novo_preco']:
            df.loc[mask, 'Preço'] = float(r['novo_preco'])
    return df


def _finaliza_df(lista_dfs, cols):
    if not lista_dfs:
        return pd.DataFrame(columns=cols)
    df_final = pd.concat(lista_dfs, ignore_index=True).drop_duplicates()
    df_final['Papel'] = df_final['Papel'].str.strip()
    return aplicar_ajustes_manuais(df_final)


def _processa_xp_rico_clear(corretora, filename, item, log, df_corretora, cell_value):
    ano_pregao = _identifica_ano_pregao(filename)
    lista_acoes = list(
        df_corretora[df_corretora['NOTA DE NEGOCIAÇÃO'].str.contains(cell_value, na=False)].index
    )

    if cell_value == "XP INVESTIMENTOS CORRETORA DE CÂMBIO, TÍTULOS E VALORES MOBILIÁRIOS S.A.":
        cell_value = 'XP INVESTIMENTOS CORRETORA DE CÂMBIO, TÍTULOS E VALORES'

    lista_bmf = list(df_corretora[df_corretora['Unnamed: 0'].str.contains(cell_value, na=False)].index)
    n1, n2 = len(lista_acoes), len(lista_bmf)

    funcao = (coir.corretoras.xp_rico_clear.xp_rico_clear if ano_pregao > 2023
              else coir.corretoras.xp_rico_clear.xp_rico_clear_old)

    if n2 >= 1:
        page_acoes = f'1-{n1}'
        page_bmf = f'{n1 + 1}-{n1 + n2}'
        return funcao(corretora, filename, item, log, page_acoes, page_bmf, control=1)
    return funcao(corretora, filename, item, log, 'all')


def _processa_xp_bmf(corretora, filename, item, log, **_):
    ano_pregao = _identifica_ano_pregao(filename)
    funcao = (coir.corretoras.xp_rico_clear_bmf.xp_rico_clear_bmf if ano_pregao > 2023
              else coir.corretoras.xp_rico_clear_bmf.xp_rico_clear_bmf_old)
    return funcao(corretora, filename, item, log, 'all', control=2)


# Tabela de despacho: (grupo, control) -> função que recebe kwargs e retorna
# (normal_df, daytrade_df, cpf, current_path)
PROCESSADORES = {
    ('XP', 1): lambda **kw: _processa_xp_rico_clear(
        kw['corretora'], kw['filename'], kw['item'], kw['log'],
        kw['df_corretora'], kw['cell_value']),
    ('XP', 2): lambda **kw: _processa_xp_bmf(
        kw['corretora'], kw['filename'], kw['item'], kw['log']),
    ('AGORA', 1): lambda **kw: coir.corretoras.agora.agora(
        kw['corretora'], kw['filename'], kw['item'], kw['log']),
    ('BTG', 1): lambda **kw: coir.corretoras.btg.btg(
        kw['corretora'], kw['filename'], kw['item'], kw['log'], 'all', control=1),
    ('BTG', 2): lambda **kw: coir.corretoras.btg_bmf.btg_bmf(
        kw['corretora'], kw['filename'], kw['item'], kw['log'], 'all', control=2),
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

    cols = ['Corretora', 'CPF', 'Conta', 'Nota', 'Data', 'C/V', 'Papel', 'Mercado',
            'Preço', 'Quantidade', 'Total', 'Custos_Fin', 'PM', 'IRRF']
    normal_dfs, daytrade_dfs = [], []
    current_path, cpf = '', ''

    for item in arquivos:
        filename = item
        log = []

        if not _valida_nota_sinacor(filename):
            print_atencao()
            print('O arquivo', '"' + basename(item).upper() + '"',
                  'NÃO é uma Nota de Corretagem no Padrão Sinacor.', '\n')
            continue

        print('processando o arquivo:', basename(item))
        log.append(datetime.today().strftime('%d/%m/%Y %H:%M:%S') +
                   ' - Processando o arquivo "' + basename(item) + '"\n')

        corretora_tabelas = tabula.read_pdf(
            filename, pandas_options={'dtype': str}, guess=False, stream=True,
            multiple_tables=True, pages='all', encoding="utf-8",
            area=(2.603, 26.609, 214.572, 561.903)
        )
        df_corretora = pd.concat(corretora_tabelas, axis=0, ignore_index=True)
        corretora = tabula.read_pdf(filename, pages='1', **kwargs,
                                     area=(2.603, 26.609, 214.572, 561.903))

        try:
            control, corretora, cell_value = valida_corretora(corretora)

            if control == 0:
                print('Corretora', cell_value, 'não implementada', '\n')
                continue

            grupo = _identifica_grupo(corretora)
            processador = PROCESSADORES.get((grupo, control))

            if processador:
                normal_df, daytrade_df, cpf, current_path = processador(
                    corretora=corretora, filename=filename, item=item, log=log,
                    df_corretora=df_corretora, cell_value=cell_value
                )
                _append_resultado(normal_df, daytrade_df, normal_dfs, daytrade_dfs)

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
                    coir.corretoras.nao_validada.nao_validada(corretora, filename, item, log, 'all')
                continue

        except ValueError as e:
            print(e)
            print('ValueError - Corretora', cell_value, 'ocorreu erro durante o processamento')
            print('das notas de corretagens', '\n')
            continue

    normal_df_final = _finaliza_df(normal_dfs, cols)
    daytrade_df_final = _finaliza_df(daytrade_dfs, cols)

    if normal_dfs or daytrade_dfs:
        coir.funcoes.arquivo_unico(current_path, cpf, normal_df_final, daytrade_df_final)

    #return normal_df_final, daytrade_df_final

# =================================================================================================
# Mensagem de alerta para os aplicativos abertos do excel
# O sistema continuará após a confirmação do usuário
# =================================================================================================
def _pergunta_sim_nao(mensagem):
    resposta = ''
    while resposta not in ('S', 'N'):
        bruto = input(mensagem).upper().strip()
        resposta = bruto[0] if bruto else ''
    return resposta

def principal():
    print()
    print('-=' * 50)
    print(f'{"SISTEMA DE CONTROLE DE OPERAÇÕES E IRPF - COIR":^100}')
    print(f'{"Versão " + __version__:^100}')
    print(f'{"Leitura, Extração e Formatação das Notas de Corretagem no padrão SINACOR":^100}')
    print('-=' * 50)
    print()
    print_atencao()
    print('Feche o Excel antes de iniciar o processamento das Notas de Corretagens.')
    print('Isso evitará erros e inconsistênica durante o processamento.\n')

    excel_fechado = _pergunta_sim_nao('O programa Excel está fechado [S/N]? ')

    if excel_fechado == 'S':
        result = subprocess.run(["taskkill", "/f", "/im", "excel.exe"],
                                 capture_output=True, text=True)
        if result.returncode == 0:
            print("Confirmado: programa Excel foi encerrado.")
        else:
            print("Confirmado: nenhum processo do Excel estava em execução.")
        print()
        print('-=' * 50)
        print('Iniciando o processamento das Notas de Corretagens...\n\n')
        extracao_nota_corretagem()

if __name__ == '__main__':
    principal()
    print('-=' * 50)
    print('Fim do processamento!', '\n')
    input('Pressione qualquer tecla para concluir.\n')
