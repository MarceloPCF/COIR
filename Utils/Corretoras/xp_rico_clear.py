# Importações padrão de bibliotecas Python
import calendar
from datetime import datetime, timedelta
from os.path import join, basename, exists

# Importações de terceiros
import tabula
import pandas as pd

# Importação das funções definidas em outro módulo
import Utils.funcoes
import Utils.Corretoras.xp_rico_clear_bmf

# =================================================================================================
# Arquivo CSV contendo o nome dos papeis no pregão da B3
# ATENÇÃO: Para cada NOVO papel, incluir o nome correto no PREGÃO e o CÓDIGO correspondente na B3
# =================================================================================================
acoes = pd.read_csv('./Utils/Tickets/acoes.csv')

# =================================================================================================
# Arquivo contendo o nome dos ativos que compoem o índice Bovespa
# =================================================================================================
opcoes = pd.read_csv('./Utils/Tickets/opcoes.csv')

# =================================================================================================
# Códigos dos ativos á vista da Bolsa de valores brasileira
# filds  = [ 'Taxa de liquidação' , 'Taxa de Registro' , 'Total CBLC' , 'Taxa de termo/opções',
# 'Taxa ANA','Emolumentos' , 'Taxa Operacional' , 'Taxa de Custódia' , 'Outros']
# Tipo mercado (OPCAO DE COMPRA | OPCAO DE VENDA | EXERC OPC VENDA | VISTA | FRACIONARIO | TERMO)
# =================================================================================================
codigos  = {'ON':'3','PN':'4','PNA':'5','PNB':'6','PNC':'7','UNT':'11'}

# =================================================================================================
# Processamento de notas de corretagens das corretoras do grupo XP (XP, Rico e Clear)
# Rotina para extração de dados no novo layout das notas de corretagens - início em 01/2024
# =================================================================================================
def xp_rico_clear(corretora,filename,item,log,page,pagebmf=0,control=0):
    # ==============================================================================================
    # Coleta de dados por área de informação - Extraindo os dados das operações na B3
    # 50.947,428.028,73.259,564.134   - Nota e data do pregão:
    # 143.172,424.894,160.278,560.256 - CPF
    # 238.028,41.348,436.198,559.868  - Informações de compra e venda
    # ==============================================================================================
    data = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False, stream=True,
    multiple_tables=True, pages=page, encoding="utf-8",
    area=(
    (50.947,428.028,73.259,564.134),
    (143.172,424.894,160.278,560.256),
    (238.028,41.348,436.198,559.868))
    )
    df = pd.concat(data,axis=0,ignore_index=True)
    df['Preço / Ajuste'] = df['Preço / Ajuste'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    df['Valor Operação / Ajuste'] = df[
    'Valor Operação / Ajuste'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    df['Quantidade'] = df['Quantidade'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    df['Nr. nota'] = df['Nr. nota'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    df['Tipo mercado'] = Utils.funcoes.sanitiza_especificacao_titulo(df['Tipo mercado'])
    df['Prazo'] = Utils.funcoes.sanitiza_especificacao_titulo(df['Prazo'])
    df['Prazo'].fillna("", inplace=True)
    df['Especificação do título'] = Utils.funcoes.sanitiza_especificacao_titulo(
    df['Especificação do título'])
    df['Obs. (*)'] = Utils.funcoes.sanitiza_observacao(df['Obs. (*)'])
    if 'Unnamed: 0' in df.columns:
        try:
            df['Unnamed: 0'] = df['Unnamed: 0'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        except:
            df['Unnamed: 0'] = pd.to_numeric(df['Unnamed: 0'], errors='coerce')
            df['Unnamed: 0'] = df['Unnamed: 0'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    if 'Unnamed: 1' in df.columns:
        try:
            df['Unnamed: 1'] = df['Unnamed: 1'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        except:
            df['Unnamed: 1'] = pd.to_numeric(df['Unnamed: 1'], errors='coerce')
            df['Unnamed: 1'] = df['Unnamed: 1'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    if 'Unnamed: 2' in df.columns:
        try:
            df['Unnamed: 2'] = df['Unnamed: 2'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        except:
            df['Unnamed: 2'] = pd.to_numeric(df['Unnamed: 2'], errors='coerce')
            df['Unnamed: 2'] = df['Unnamed: 2'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    # =============================================================================================
    # Coleta de dados por área de informação - Extraindo os dados das taxas e impostos
    # 49.543,428.003,68.913,562.103 - nota e data do pregão:
    # 439.178,33.898,617.233,546.458 - Resumo dos negócios, Resumo financeiro e Custos operacionais
    # =============================================================================================
    data = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False, stream=True,
    multiple_tables=True, pages='all', encoding="utf-8",
    area=(
    (49.543,428.003,68.913,562.103),
    (435.453,32.408,617.978,546.458))
    )
    df_gastos = pd.concat(data,axis=0,ignore_index=True)
    df_gastos['Nr. nota'] = df_gastos['Nr. nota'].apply(
    Utils.funcoes.sanitiza_moeda).astype('float')
    if 'Unnamed: 0' in df_gastos.columns:
        df_gastos['Unnamed: 0'] = df_gastos['Unnamed: 0'].apply(
        Utils.funcoes.sanitiza_moeda).astype('float')
        df_gastos['Unnamed: 0'].fillna(0, inplace=True)
    if 'Unnamed: 1' in df_gastos.columns:
        df_gastos['Unnamed: 1'] = df_gastos['Unnamed: 1'].apply(
        Utils.funcoes.sanitiza_moeda).astype('float')
        df_gastos['Unnamed: 1'].fillna(0, inplace=True)
    lista = list(
    df_gastos[df_gastos['Resumo dos Negócios'].str.contains("Valor das operações",na=False)].index)
    note_taxa = []

    # Obtem o número da conta na corretora
    conta = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False, stream=True,
    multiple_tables=True, pages='1', encoding="utf-8", area=(156.823,429.493,177.683,522.617))
    conta = pd.concat(conta,axis=0,ignore_index=True)
    conta = conta['Unnamed: 0'].iloc[0].strip().lstrip('0')

    # Verifica se a Nota de Corretagem já foi processada anteriormente
    #current_path = './Resultado/'
    #cpf = str(df['C.P.F./C.N.P.J/C.V.M./C.O.B.'][1])
    #nome = conta + '_' + df_gastos['Data pregão'][0][6:10]
    #nome += '_' + df_gastos['Data pregão'][0][3:5] + '.xlsx'
    #folder_prefix = str(
    #df['C.P.F./C.N.P.J/C.V.M./C.O.B.'][1] +'/'+ corretora +'/'+ df_gastos['Data pregão'][0][6:10])
    #folder_path = join(current_path, folder_prefix)
    #if exists(folder_path+'/'+nome):
    #    log.append(Utils.funcoes.verifica_nota_corretagem(folder_path,nome,item))
    #    Utils.funcoes.log_processamento(current_path,cpf,log)
    #    return

    for current_row in lista:
        nota = df_gastos['Nr. nota'].iloc[current_row-8]
        data = datetime.strptime(df_gastos['Data pregão'].iloc[current_row-8], '%d/%m/%Y').date()
        total = df_gastos['Unnamed: 0'].iloc[current_row]
        vendas = df_gastos['Unnamed: 0'].iloc[current_row-6]
        liquidacao = df_gastos['Unnamed: 1'].iloc[current_row-5]
        registro = df_gastos['Unnamed: 1'].iloc[current_row-4]
        emolumentos = df_gastos['Unnamed: 1'].iloc[current_row+1]
        corretagem = df_gastos['Unnamed: 1'].iloc[current_row+5]
        imposto = df_gastos['Unnamed: 1'].iloc[current_row+8]
        irrf = df_gastos['Unnamed: 1'].iloc[current_row+9]
        outros = df_gastos['Unnamed: 1'].iloc[current_row+10]
        ir_daytrade = str(df_gastos['Resumo dos Negócios'].iloc[current_row+10])
        if ir_daytrade != "nan":
            ir_daytrade = ir_daytrade.split("Projeção R$ ")[1]
            outros = df_gastos['Unnamed: 1'].iloc[current_row+11]
        else:
            ir_daytrade = "0"
        ir_daytrade = float(ir_daytrade.replace('.','').replace(',','.'))
        basecalculo = str(df_gastos['Resumo Financeiro'].iloc[current_row+9])
        if basecalculo != "nan":
            basecalculo = basecalculo.split("base R$")[1]
            if basecalculo == "":
                basecalculo = "0"
        else:
            basecalculo = "0"
        basecalculo = float(basecalculo.replace('.','').replace(',','.'))
        row_data = [
        nota,data,total,vendas,liquidacao,registro,emolumentos,corretagem,imposto,outros,
        emolumentos+liquidacao+registro,corretagem+imposto+outros,irrf,ir_daytrade,basecalculo
        ]
        note_taxa.append(row_data)

    cols = ['Nota','Data','Total','Vendas','Liquidação','Registro','Emolumentos','Corretagem',
    'Imposto','Outros','Custos_Fin','Custos_Op','IRRF','IR_DT','BaseCalculo']
    taxas_df = pd.DataFrame(data=note_taxa, columns=cols)
    taxas_df = taxas_df.drop_duplicates(subset='Nota', keep='last', ignore_index=True)
    cont_notas = len(taxas_df['Nota'])
    if cont_notas > 1:
        log.append(
        'Foram processadas ' + str(cont_notas) + ' notas de corretagens do mercado à vista.\n')
    else:
        log.append(
        'Foi processada ' + str(cont_notas) + ' nota de corretagem do mercado à vista.\n')

    # Incluir aqui a etapa para obter lista de linhas de cada operação
    operacoes = list(df[df['Negociação'].str.contains("1-BOVESPA",na=False)].index)
    current_path = './Resultado/'
    note_data = []
    numero_nota = 0
    cpf = ''
    #nome = ''
    ano = ''
    temp = ''
    for current_row in operacoes:
        cell_value = df['Nr. nota'].iloc[current_row-2]
        if cell_value > 0:
            numero_nota = df['Nr. nota'].iloc[current_row-2]
            data = df['Data pregão'].iloc[current_row-2]
            if ano == '':
                cpf = df['C.P.F./C.N.P.J/C.V.M./C.O.B.'].iloc[current_row-1]
                #nome = conta + '_' + data[6:10] + '_' + data[3:5] + '.xlsx'
                ano = data[6:10]
            data = datetime.strptime(df['Data pregão'].iloc[current_row-2], '%d/%m/%Y').date()

        # Tipo de operação (Compra ou Venda)
        c_v = df['C/V'].iloc[current_row].strip()

        # Nome do ativo no pregão
        stock_title = df['Especificação do título'].iloc[current_row].strip()

        operacao = df['Obs. (*)'].iloc[current_row]
        if operacao == "D":
            operacao = "DayTrade"
        else:
            operacao = "Normal"

        # Tipo de Mercado operado
        if df['Tipo mercado'].iloc[current_row] == "VISTA":
            mercado = "VISTA"
        elif df['Tipo mercado'].iloc[current_row] == "OPCAO DE COMPRA":
            mercado = "CALL"
        elif df['Tipo mercado'].iloc[current_row] == "OPCAO DE VENDA":
            mercado = "PUT"
        elif df['Tipo mercado'].iloc[current_row] == "EXERC OPC VENDA":
            mercado = "EXERC PUT"
        elif df['Tipo mercado'].iloc[current_row] == "EXERC OPC COMPRA":
            mercado = "EXERC CALL"
        else:
            mercado = df['Tipo mercado'].iloc[current_row]

        # Prazo de Vencimento da Opção
        if df['Prazo'].iloc[current_row] != "" and mercado != "VISTA" and df['Prazo'].iloc[current_row] == str:
            monthrange = calendar.monthrange(
            2000 + int(df['Prazo'].iloc[current_row][3:]),int(df['Prazo'].iloc[current_row][0:2])
            )
            prazo = str(monthrange[1]) +'/'+ str(int(df['Prazo'].iloc[current_row][0:2])) +'/'+str(
            2000 + int(df['Prazo'].iloc[current_row][3:])
            )
            prazo = datetime.strptime(prazo, '%d/%m/%Y').date()
        elif df['Prazo'].iloc[current_row] and df['Tipo mercado'].iloc[current_row].lower() == "termo":
            prazo = int(df['Prazo'].iloc[current_row])
        else:
            prazo = ""

        # Exercicio de opção de compra/venda
        if df['Tipo mercado'].iloc[current_row].split(" ")[0] == "EXERC":
            exercicio = df['Especificação do título'].iloc[current_row]
            exercicio = exercicio[:-1]
        elif df['Tipo mercado'].iloc[current_row].split(" ")[0] == "OPCAO":
            exercicio = Utils.funcoes.converte_opcao_ticket(
            df['Especificação do título'].iloc[current_row])
        else:
            exercicio = ""

        # Altera o número de dias de um contrato a Termo para a
        # data de vencimento desse contrato
        if mercado in "TERMOTermoTERMO":
            prazo = data + timedelta(days=prazo)

        # Quantidade operada de cada ativo por nota de corretagem
        quantidade = Utils.funcoes.quantidade_operada(
        df['Quantidade'].iloc[current_row],
        df['Unnamed: 0'].iloc[current_row] if 'Unnamed: 0' in df.columns else 0,
        df['Unnamed: 1'].iloc[current_row] if 'Unnamed: 1' in df.columns else 0,
        df['Unnamed: 2'].iloc[current_row] if 'Unnamed: 2' in df.columns else 0
        )

        # Valor total de cada operação por nota de corretagem
        valor_total = Utils.funcoes.valor_total_ativo(
        df['Valor Operação / Ajuste'].iloc[current_row],
        df['Unnamed: 2'].iloc[current_row] if 'Unnamed: 2' in df.columns else 0,
        df['Unnamed: 1'].iloc[current_row] if 'Unnamed: 1' in df.columns else 0
        )

        # Preço unitário da operação de cada ativo por nota de corretagem
        preco_unitario = valor_total / quantidade

        # Dividindo os custos e o IRRF por operação
        custo_financeiro,irrf_operacao = Utils.funcoes.custos_por_operacao(taxas_df, numero_nota,
        c_v, valor_total, operacao)

        # Susbstitui o nome do papel no pregão pelo seu respectivo código na B3 no padrão "XXXX3"
        # Caso seja uma opção de compra/venda o código continuará o mesmo
        if df['Tipo mercado'].iloc[current_row].lower() in ["vista", "fracionario", "termo"]:
            stock_title,log_nome_pregao = Utils.funcoes.nome_pregao(acoes, stock_title, data)
            if log_nome_pregao != temp:
                temp = log_nome_pregao
                log.append(log_nome_pregao)
        elif df['Tipo mercado'].iloc[current_row].split(" ")[0] == "EXERC":
            stock_title,log_nome_pregao = Utils.funcoes.nome_pregao_opcoes(opcoes,stock_title,data)
            if log_nome_pregao != temp:
                temp = log_nome_pregao
                log.append(log_nome_pregao)
        elif df['Tipo mercado'].iloc[current_row].split(" ")[0] == "OPCAO":
            stock_title = stock_title.split(" ")[0]
        else:
            stock_title,log_nome_pregao = Utils.funcoes.nome_pregao(acoes, stock_title, data)
            if log_nome_pregao != temp:
                temp = log_nome_pregao
                log.append(log_nome_pregao)

        # Calculando o preço médio de cada operação
        pm = Utils.funcoes.preco_medio(c_v,valor_total,custo_financeiro,quantidade)

        row_data = [
        corretora, cpf, numero_nota, data, c_v, stock_title, operacao, preco_unitario, quantidade,
        valor_total,custo_financeiro, pm, irrf_operacao,mercado,prazo,exercicio
        ]
        note_data.append(row_data)
    cols = ['Corretora','CPF','Nota','Data','C/V','Papel','Operacao','Preço','Quantidade','Total',
    'Custos_Fin', 'PM', 'IRRF', 'Mercado', 'Prazo', 'Exercicio'
    ]
    note_df = pd.DataFrame(data=note_data, columns=cols)

    # Agrupar os dados de preço e quantidade por cada ativo C/V em cada nota de corretagem
    grouped = Utils.funcoes.agrupar(note_df)
    grouped = grouped[cols]

    # Seleção de papel isento de IR (IRRF e IRPF).
    # Apenas uma operação (um papel) está sendo analisada por NC
    note_data,log_isecao = Utils.funcoes.isencao_imposto_renda(taxas_df,grouped,note_data)
    log.append(log_isecao)
    cols = ['Corretora','CPF','Nota','Data','C/V','Papel','Operacao','Preço','Quantidade','Total',
    'Custos_Fin', 'PM', 'IRRF', 'Mercado', 'Prazo', 'Exercicio']
    note_df = pd.DataFrame(data=note_data, columns=cols)

    # Atulizar os dados de preço e quantidade com a correção de C/V a maior no DayTrade
    grouped = Utils.funcoes.agrupar(note_df)
    grouped = grouped[cols]

    # Agrupa as operações por tipo de operação (Normal ou Daytrade)
    try:
        normal_df,daytrade_df,result = Utils.funcoes.agrupar_operacoes(grouped,cols)
        # Insere o valor do IR para as operações de Daytrade"
        note_data,taxas_df,log_daytrade_ir = Utils.funcoes.daytrade_ir(
        result,taxas_df,note_data,grouped)
        log.append(log_daytrade_ir)
        cols = ['Corretora','CPF', 'Nota', 'Data', 'C/V', 'Papel', 'Operacao','Preço','Quantidade',
        'Total','Custos_Fin', 'PM', 'IRRF','Mercado','Prazo','Exercicio'
        ]
        note_df = pd.DataFrame(data=note_data, columns=cols)
        # Refaz o agrupamento para atualizar os dados de preço e quantidade por cada ativo C/V.
        grouped = Utils.funcoes.agrupar(note_df)
        grouped = grouped[cols]
    except ValueError:
        Utils.funcoes.agrupar_operacoes(grouped,cols)
        #normal_df,daytrade_df = Utils.funcoes.agrupar_operacoes(grouped,cols)

    # Acrescentando os custos operacionais (Corretagem, Imposto e Outros)
    grouped = Utils.funcoes.custos_operacionais(grouped,taxas_df)

    # Obtendo o valor correto do preço unitário de cada operação
    grouped['Preço'] = grouped['Total'] / grouped['Quantidade']

    # Obtendo o valor correto do preço médio de cada operação
    grouped['PM'] = Utils.funcoes.preco_medio_correcao(grouped)

    # Inseri o número da conta na corretora
    grouped.insert(2,"Conta",int(conta),True)
    taxas_df.insert(0,"Conta",int(conta),True)

    # Agrupa as operações por tipo de trade com correção de compra/venda a maior no DayTrade
    normal_df,daytrade_df = Utils.funcoes.agrupar_operacoes_correcao(grouped,cols)
    cols = ['Corretora','CPF','Conta','Nota','Data','C/V','Papel','Mercado','Preço','Quantidade',
    'Total','Custos_Fin', 'PM', 'IRRF','Prazo','Exercicio'
    ]
    if not normal_df.empty:
        normal_df = normal_df[cols]
    if not daytrade_df.empty:
        daytrade_df = daytrade_df[cols]

    # Cria o caminho completo de pastas/subpasta para salvar o resultado do processamento
    #current_path = './Resultado/'
    #folder_prefix = cpf+'/'+corretora+'/'+ano
    #folder_path = join(current_path, folder_prefix)
    #log_move_resultado,pagebmf = Utils.funcoes.move_resultado(folder_path,cpf,nome,item,pagebmf)
    log_move_resultado = Utils.funcoes.move_resultado(cpf)
    log.append(log_move_resultado)

    # nome_arquivo = cpf +'/'+ cpf+".xlsb"
    # caminho_arquivo = join(current_path, nome_arquivo)
    # Utils.funcoes.ler_dados_excel(current_path,cpf)

    # Disponibiliza os dados coletados em um arquivo .xlsx separado por mês
    # Utils.funcoes.arquivo_separado(folder_path,nome,note_df,normal_df,daytrade_df,taxas_df)

    # Disponibiliza todos os dados coletados de todos os arquivos processados em um único arquivo
    # Utils.funcoes.arquivo_unico(current_path,cpf,note_df,normal_df,daytrade_df,taxas_df)
    Utils.funcoes.arquivo_unico(current_path,cpf,normal_df,daytrade_df)

    # Disponibiliza todos os dados coletados de todos os arquivos processados em um único arquivo
    if pagebmf != 0 and control != 0:
        Utils.Corretoras.xp_rico_clear_bmf.xp_rico_clear_bmf(
        corretora,filename,item,log,pagebmf,control
        )

    # Cria o caminho completo de pastas/subpastas para mover os arquivos já processados.
    log_move_saida = Utils.funcoes.move_saida(cpf,corretora,ano,item)
    log.append(log_move_saida)

    # Cria um arquivo de LOG para armazenar os dados do processamento
    log.append(
    'As Notas de Corretagem do arquivo "'+basename(item)+'" foram processadas com sucesso.\n'
    )
    Utils.funcoes.log_processamento(current_path,cpf,log)


# =================================================================================================
# Processamento de notas de corretagens das corretoras do grupo XP (XP, Rico e Clear)
# Rotina para extração de dados de notas de corretagens até 12/2023
# =================================================================================================
def xp_rico_clear_old(corretora,filename,item,log,page,pagebmf=0,control=0):
    # =============================================================================================
    # Coleta de dados por área de informação - Extraindo os dados das operações na B3
    # =============================================================================================
    # 50.947,428.028,73.259,564.134   - Nota e data do pregão:
    # 143.172,424.894,160.278,560.256 - CPF
    # 240.603,32.194,448.109,561.0    - Informações de compra e venda:
    data = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False, stream=True,
    multiple_tables=True, pages=page, encoding="utf-8",
    area=(
    (50.947,428.028,73.259,564.134),
    (143.172,424.894,160.278,560.256),
    (240.603,32.194,448.109,561.0))
    )
    df = pd.concat(data,axis=0,ignore_index=True)
    df['Preço / Ajuste'] = df['Preço / Ajuste'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    df['Valor Operação / Ajuste'] = df['Valor Operação / Ajuste'].apply(
    Utils.funcoes.sanitiza_moeda).astype('float')
    df['Quantidade'] = df['Quantidade'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    df['Nr. nota'] = df['Nr. nota'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    df['Tipo mercado'] = Utils.funcoes.sanitiza_especificacao_titulo(df['Tipo mercado'])
    df['Prazo'] = Utils.funcoes.sanitiza_especificacao_titulo(df['Prazo'])
    df['Prazo'].fillna("", inplace=True)
    df['Especificação do título'] = Utils.funcoes.sanitiza_especificacao_titulo(
    df['Especificação do título'])
    df['Obs. (*)'] = Utils.funcoes.sanitiza_observacao(df['Obs. (*)'])
    if 'Unnamed: 0' in df.columns:
        try:
            df['Unnamed: 0'] = df['Unnamed: 0'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        except:
            df['Unnamed: 0'] = pd.to_numeric(df['Unnamed: 0'], errors='coerce')
            df['Unnamed: 0'] = df['Unnamed: 0'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    if 'Unnamed: 1' in df.columns:
        try:
            df['Unnamed: 1'] = df['Unnamed: 1'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        except:
            df['Unnamed: 1'] = pd.to_numeric(df['Unnamed: 1'], errors='coerce')
            df['Unnamed: 1'] = df['Unnamed: 1'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    if 'Unnamed: 2' in df.columns:
        try:
            df['Unnamed: 2'] = df['Unnamed: 2'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        except:
            df['Unnamed: 2'] = pd.to_numeric(df['Unnamed: 2'], errors='coerce')
            df['Unnamed: 2'] = df['Unnamed: 2'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    # =============================================================================================
    # Coleta de dados por área de informação - Extraindo os dados das taxas e impostos
    # 50.947,428.028,73.259,564.134  - nota e data do pregão:
    # 450.341,32.576,639.253,544.276 -  Resumo dos negócios e financeiro e Custos operacionais:
    # =============================================================================================
    data = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False, stream=True,
    multiple_tables=True, pages='all', encoding="utf-8",
    area=(
    (53.178,428.995,71.772,561.382),
    (450.341,32.576,639.253,544.276))
    )
    df_gastos = pd.concat(data,axis=0,ignore_index=True)
    df_gastos['Nr. nota'] = df_gastos['Nr. nota'].apply(
    Utils.funcoes.sanitiza_moeda).astype('float')
    if 'Unnamed: 0' in df_gastos.columns:
        df_gastos['Unnamed: 0'] = df_gastos[
        'Unnamed: 0'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        df_gastos['Unnamed: 0'].fillna(0, inplace=True)
    if 'Unnamed: 1' in df_gastos.columns:
        df_gastos['Unnamed: 1'] = df_gastos[
        'Unnamed: 1'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        df_gastos['Unnamed: 1'].fillna(0, inplace=True)
    lista = list(
    df_gastos[df_gastos['Resumo dos Negócios'].str.contains("Valor das operações",na=False)].index)
    note_taxa = []

    # Obtem o número da conta na corretora
    conta = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False, stream=True,
    multiple_tables=True, pages='1', encoding="utf-8", area=(160.278,426.541,179.616,520.253))
    conta = pd.concat(conta,axis=0,ignore_index=True)
    conta = conta['Unnamed: 0'].iloc[0].strip().lstrip('0')

    # Verifica se a Nota de Corretagem já foi processada anteriormente
    #current_path = './Resultado/'
    #cpf = str(df['C.P.F./C.N.P.J/C.V.M./C.O.B.'][1])
    #nome = conta + '_' + df_gastos['Data pregão'][0][6:10]
    #nome +=  '_' + df_gastos['Data pregão'][0][3:5] + '.xlsx'
    #current_path = './Resultado/'
    #folder_prefix = str(
    #df['C.P.F./C.N.P.J/C.V.M./C.O.B.'][1] +'/'+ corretora +'/'+ df_gastos['Data pregão'][0][6:10])
    #folder_path = join(current_path, folder_prefix)
    #if exists(folder_path+'/'+nome):
    #    log.append(Utils.funcoes.verifica_nota_corretagem(folder_path,nome,item))
    #    Utils.funcoes.log_processamento(current_path,cpf,log)
    #    return

    for current_row in lista:
        nota = df_gastos['Nr. nota'].iloc[current_row-8]
        data = datetime.strptime(df_gastos['Data pregão'].iloc[current_row-8], '%d/%m/%Y').date()
        total = df_gastos['Unnamed: 0'].iloc[current_row]
        vendas = df_gastos['Unnamed: 0'].iloc[current_row-6]
        liquidacao = df_gastos['Unnamed: 1'].iloc[current_row-5]
        registro = df_gastos['Unnamed: 1'].iloc[current_row-4]
        emolumentos = df_gastos['Unnamed: 1'].iloc[current_row+1]
        corretagem = df_gastos['Unnamed: 1'].iloc[current_row+5]
        imposto = df_gastos['Unnamed: 1'].iloc[current_row+8]
        irrf = df_gastos['Unnamed: 1'].iloc[current_row+9]
        outros = df_gastos['Unnamed: 1'].iloc[current_row+10]
        ir_daytrade = str(df_gastos['Resumo dos Negócios'].iloc[current_row+10])
        if ir_daytrade != "nan":
            ir_daytrade = ir_daytrade.split("Projeção R$ ")[1]
            outros = df_gastos['Unnamed: 1'].iloc[current_row+11]
        else:
            ir_daytrade = "0"
        ir_daytrade = float(ir_daytrade.replace('.','').replace(',','.'))
        basecalculo = str(df_gastos['Resumo Financeiro'].iloc[current_row+9])
        if basecalculo != "nan":
            basecalculo = basecalculo.split("base R$")[1]
        else:
            basecalculo = "0"
        basecalculo = float(basecalculo.replace('.','').replace(',','.'))
        row_data = [
        nota,data,total,vendas,liquidacao,registro,emolumentos,corretagem,imposto,outros,
        emolumentos+liquidacao+registro,corretagem+imposto+outros,irrf,ir_daytrade,basecalculo]
        note_taxa.append(row_data)
    cols = ['Nota','Data','Total','Vendas','Liquidação','Registro','Emolumentos','Corretagem',
    'Imposto','Outros','Custos_Fin','Custos_Op','IRRF','IR_DT','BaseCalculo']
    taxas_df = pd.DataFrame(data=note_taxa, columns=cols)
    taxas_df = taxas_df.drop_duplicates(subset='Nota', keep='last', ignore_index=True)
    cont_notas = len(taxas_df['Nota'])
    if cont_notas > 1:
        log.append(
        'Serão processadas ' + str(cont_notas) + ' notas de corretagens do mercado à vista.\n')
    else:
        log.append(
        'Será processada ' + str(cont_notas) + ' nota de corretagem do mercado à vista.\n')

    # Incluir aqui a etapa para obter lista de linhas de cada operação
    operacoes = list(df[df['Negociação'].str.contains("1-BOVESPA",na=False)].index)
    current_path = './Resultado/'
    note_data = []
    numero_nota = 0
    cpf = ''
    #nome = ''
    ano = ''
    temp = ''
    for current_row in operacoes:
        cell_value = df['Nr. nota'].iloc[current_row-2]
        if cell_value > 0:
            numero_nota = df['Nr. nota'].iloc[current_row-2]
            data = df['Data pregão'].iloc[current_row-2]
            if ano == '':
                cpf = df['C.P.F./C.N.P.J/C.V.M./C.O.B.'].iloc[current_row-1]
                #nome = conta + '_' + data[6:10] + '_' + data[3:5] + '.xlsx'
                ano = data[6:10]
            data = datetime.strptime(df['Data pregão'].iloc[current_row-2], '%d/%m/%Y').date()

        # Tipo de operação (Compra ou Venda)
        c_v = df['C/V'].iloc[current_row].strip()

        # Nome do ativo no pregão
        stock_title = df['Especificação do título'].iloc[current_row].strip()

        operacao = df['Obs. (*)'].iloc[current_row]
        if operacao == "D":
            operacao = "DayTrade"
        else:
            operacao = "Normal"

        # Tipo de Mercado operado
        if df['Tipo mercado'].iloc[current_row] == "VISTA":
            mercado = "VISTA"
        elif df['Tipo mercado'].iloc[current_row] == "OPCAO DE COMPRA":
            mercado = "CALL"
        elif df['Tipo mercado'].iloc[current_row] == "OPCAO DE VENDA":
            mercado = "PUT"
        elif df['Tipo mercado'].iloc[current_row] == "EXERC OPC VENDA":
            mercado = "EXERC PUT"
        elif df['Tipo mercado'].iloc[current_row] == "EXERC OPC COMPRA":
            mercado = "EXERC CALL"
        else:
            mercado = df['Tipo mercado'].iloc[current_row]

        # Prazo de Vencimento da Opção
        if df['Prazo'].iloc[current_row] != "" and mercado != "VISTA" and df['Prazo'].iloc[current_row] == str:
            monthrange = calendar.monthrange(
            2000 + int(df['Prazo'].iloc[current_row][3:]),int(df['Prazo'].iloc[current_row][0:2]))
            prazo = str(monthrange[1]) +'/'+ str(int(df['Prazo'].iloc[current_row][0:2])) +'/'+str(
            2000 + int(df['Prazo'].iloc[current_row][3:]))
            prazo = datetime.strptime(prazo, '%d/%m/%Y').date()
        elif df['Prazo'].iloc[current_row] != "" and df['Tipo mercado'].iloc[current_row].lower() == "termo":
            prazo = int(df['Prazo'].iloc[current_row])
        else:
            prazo = ""

        # Exercicio de opção de compra/venda
        if df['Tipo mercado'].iloc[current_row].split(" ")[0] == "EXERC":
            exercicio = df['Especificação do título'].iloc[current_row]
            exercicio = exercicio[:-1]
        elif df['Tipo mercado'].iloc[current_row].split(" ")[0] == "OPCAO":
            exercicio = Utils.funcoes.converte_opcao_ticket(
            df['Especificação do título'].iloc[current_row])
        else:
            exercicio = ""

        # Altera o número de dias de um contrato a Termo para a
        # data de vencimento desse contrato
        if mercado in "TERMOTermoTERMO":
            prazo = data + timedelta(days=prazo)

        # Quantidade operada de cada ativo por nota de corretagem
        quantidade = Utils.funcoes.quantidade_operada(
        df['Quantidade'].iloc[current_row],
        df['Unnamed: 0'].iloc[current_row] if 'Unnamed: 0' in df.columns else 0,
        df['Unnamed: 1'].iloc[current_row] if 'Unnamed: 1' in df.columns else 0,
        df['Unnamed: 2'].iloc[current_row] if 'Unnamed: 2' in df.columns else 0
        )

        # Valor total de cada operação por nota de corretagem
        valor_total = Utils.funcoes.valor_total_ativo(
        df['Valor Operação / Ajuste'].iloc[current_row],
        df['Unnamed: 2'].iloc[current_row] if 'Unnamed: 2' in df.columns else 0,
        df['Unnamed: 1'].iloc[current_row] if 'Unnamed: 1' in df.columns else 0
        )

        # Preço unitário da operação de cada ativo por nota de corretagem
        preco_unitario = valor_total / quantidade

        # Dividindo os custos e o IRRF por operação
        custo_financeiro,irrf_operacao = Utils.funcoes.custos_por_operacao(
        taxas_df,numero_nota,c_v,valor_total,operacao
        )

        # Susbstitui o nome do papel no pregão pelo seu respectivo código na B3 no padrão "XXXX3"
        # Caso seja uma opção de compra/venda o código continuará o mesmo
        if df['Tipo mercado'].iloc[current_row].lower() in ["vista","fracionario","termo"]:
            stock_title,log_nome_pregao = Utils.funcoes.nome_pregao(acoes, stock_title, data)
            if log_nome_pregao != temp:
                temp = log_nome_pregao
                log.append(log_nome_pregao)
        elif df['Tipo mercado'].iloc[current_row].split(" ")[0] == "EXERC":
            stock_title,log_nome_pregao = Utils.funcoes.nome_pregao_opcoes(
            opcoes, stock_title, data)
            if log_nome_pregao != temp:
                temp = log_nome_pregao
                log.append(log_nome_pregao)
        elif df['Tipo mercado'].iloc[current_row].split(" ")[0] == "OPCAO":
            stock_title = stock_title.split(" ")[0]
        else:
            stock_title,log_nome_pregao = Utils.funcoes.nome_pregao(acoes, stock_title, data)
            if log_nome_pregao != temp:
                temp = log_nome_pregao
                log.append(log_nome_pregao)

        # Calculando o preço médio de cada operação
        pm = Utils.funcoes.preco_medio(c_v,valor_total,custo_financeiro,quantidade)

        row_data = [
        corretora, cpf, numero_nota, data, c_v, stock_title, operacao, preco_unitario,
        quantidade, valor_total,custo_financeiro, pm, irrf_operacao,mercado,prazo,exercicio
        ]
        note_data.append(row_data)
    cols = [
    'Corretora','CPF', 'Nota', 'Data', 'C/V', 'Papel', 'Operacao','Preço', 'Quantidade',
    'Total', 'Custos_Fin', 'PM','IRRF', 'Mercado', 'Prazo', 'Exercicio'
    ]
    note_df = pd.DataFrame(data=note_data, columns=cols)

    # Agrupar os dados de preço e quantidade por cada ativo C/V em cada nota de corretagem
    grouped = Utils.funcoes.agrupar(note_df)
    grouped = grouped[cols]

    # Seleção de papel isento de IR (IRRF e IRPF).
    # Apenas uma operação (um papel) está sendo analisada por NC
    note_data,log_isecao = Utils.funcoes.isencao_imposto_renda(taxas_df,grouped,note_data)
    log.append(log_isecao)
    cols = [
    'Corretora','CPF','Nota','Data','C/V','Papel','Operacao','Preço','Quantidade','Total',
    'Custos_Fin','PM','IRRF', 'Mercado', 'Prazo', 'Exercicio'
    ]
    note_df = pd.DataFrame(data=note_data, columns=cols)

    # Atulizar os dados de preço e quantidade com a correção de compra/venda a maior no DayTrade
    grouped = Utils.funcoes.agrupar(note_df)
    grouped = grouped[cols]

    # Agrupa as operações por tipo de operação (Normal ou Daytrade)
    try:
        normal_df,daytrade_df,result = Utils.funcoes.agrupar_operacoes(grouped,cols)
        # Insere o valor do IR para as operações de Daytrade"
        note_data,taxas_df,log_daytrade_ir = Utils.funcoes.daytrade_ir(
        result,taxas_df,note_data,grouped)
        log.append(log_daytrade_ir)
        cols = [
        'Corretora','CPF','Nota','Data','C/V','Papel','Operacao','Preço','Quantidade',
        'Total', 'Custos_Fin','PM', 'IRRF','Mercado','Prazo','Exercicio'
        ]
        note_df = pd.DataFrame(data=note_data, columns=cols)
        # Atualizar os dados de preço e quantidade por cada ativo comprado/vendido
        grouped = Utils.funcoes.agrupar(note_df)
        grouped = grouped[cols]
    except ValueError:
        Utils.funcoes.agrupar_operacoes(grouped,cols)
        #normal_df,daytrade_df = Utils.funcoes.agrupar_operacoes(grouped,cols)
        #excluir esse retorno e testar!!!

    # Acrescentando os custos operacionais (Corretagem, Imposto e Outros)
    grouped = Utils.funcoes.custos_operacionais(grouped,taxas_df)

    # Obtendo o valor correto do preço unitário de cada operação
    grouped['Preço'] = grouped['Total'] / grouped['Quantidade']

    # Obtendo o valor correto do preço médio de cada operação
    grouped['PM'] = Utils.funcoes.preco_medio_correcao(grouped)

    # Inseri o número da conta na corretora
    grouped.insert(2,"Conta",int(conta),True)
    taxas_df.insert(0,"Conta",int(conta),True)

    # Agrupa as operações por tipo de trade com correção de compra/venda a maior no DayTrade
    normal_df,daytrade_df = Utils.funcoes.agrupar_operacoes_correcao(grouped,cols)
    cols = [
    'Corretora','CPF','Conta','Nota','Data', 'C/V', 'Papel','Mercado','Preço','Quantidade',
    'Total', 'Custos_Fin','PM', 'IRRF','Prazo','Exercicio'
    ]
    if not normal_df.empty:
        normal_df = normal_df[cols]
    if not daytrade_df.empty:
        daytrade_df = daytrade_df[cols]

    # Cria o caminho completo de pastas/subpasta para salvar o resultado do processamento
    #current_path = './Resultado/'
    # folder_prefix = cpf+'/'+corretora+'/'+ano
    # folder_path = join(current_path, folder_prefix)
    log_move_resultado = Utils.funcoes.move_resultado(cpf)
    log.append(log_move_resultado)

    # Disponibiliza os dados coletados em um arquivo .xlsx separado por mês
    # Utils.funcoes.arquivo_separado(folder_path,nome,note_df,normal_df,daytrade_df,taxas_df)

    # Disponibiliza todos os dados coletados de todos os arquivos processados em um único arquivo
    # Utils.funcoes.arquivo_unico(current_path,cpf,note_df,normal_df,daytrade_df,taxas_df)
    Utils.funcoes.arquivo_unico(current_path,cpf,normal_df,daytrade_df)

    # Disponibiliza todos os dados coletados de todos os arquivos processados em um único arquivo
    if pagebmf != 0 and control != 0:
        Utils.Corretoras.xp_rico_clear_bmf.xp_rico_clear_bmf_old(
        corretora,filename,item,log,pagebmf,control)

    # Cria o caminho completo de pastas/subpastas para mover os arquivos já processados.
    log_move_saida = Utils.funcoes.move_saida(cpf,corretora,ano,item)
    log.append(log_move_saida)

    # Cria um arquivo de LOG para armazenar os dados do processamento
    log.append(
    'As Notas de Corretagem do arquivo "'+basename(item)+'" foram processadas com sucesso.\n')
    Utils.funcoes.log_processamento(current_path,cpf,log)
