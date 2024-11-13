# Importações padrão de bibliotecas Python
import calendar
from datetime import datetime
from os.path import join, basename, exists

# Importações de terceiros
import tabula
import pandas as pd

# Importação das funções definidas em outro módulo
import Utils.funcoes
import Utils.Corretoras.btg_bmf

# ==================================================================================================
# Arquivo CSV contendo o nome dos papeis no pregão da B3
# ATENÇÃO: Para cada NOVO papel, incluir o nome correto no PREGÃO e o CÓDIGO correspondente na B3
# ==================================================================================================
acoes = pd.read_csv('./Utils/Tickets/acoes.csv')

# ==================================================================================================
# Arquivo contendo o nome dos ativos que compoem o índice Bovespa
# ==================================================================================================
opcoes = pd.read_csv('./Utils/Tickets/opcoes.csv')

# ==================================================================================================
# Códigos dos ativos á vista da Bolsa de valores brasileira
# filds  = ['Taxa de liquidação' , 'Taxa de Registro' , 'Total CBLC' , 'Taxa de termo/opções',
# 'Taxa ANA' , 'Emolumentos' , 'Taxa Operacional' , 'Taxa de Custódia' , 'Outros']
# Tipo mercado (OPCAO DE COMPRA|OPCAO DE VENDA|EXERC OPC VENDA|VISTA|FRACIONARIO|TERMO)
# ==================================================================================================
codigos  = { 'ON' : '3' , 'PN' : '4' , 'PNA' : '5' , 'PNB' : '6' , 'PNC' : '7' , 'UNT' : '11' }

# ==================================================================================================
# Processamento de notas de corretagens das corretoras do grupo BTG Pactual
# ==================================================================================================
def btg(corretora,filename,item,log,page,pagebmf=0,control=0):
    # ==============================================================================================
    # Coleta de dados por área de informação - Extraindo os dados das operações na B3
    # 51.691,400.509,72.516,568.597   - Nota e data do pregão:
    # 141.684,423.566,158.791,567.109 - CPF
    # 240.603,24.172,497.941,570.828  - Informações de compra e venda:
    # ==============================================================================================
    data = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False, stream=True,
    multiple_tables=True, pages=page, encoding="utf-8",
    area=((51.691,400.509,72.516,568.597),
    (141.684,423.566,158.791,567.109),
    (240.603,24.172,497.941,570.828))
    )
    df = pd.concat(data,axis=0,ignore_index=True)
    df['Preço / Ajuste'] = df['Preço / Ajuste'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    df['Valor Operação / Ajuste'] = df[
    'Valor Operação / Ajuste'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    df['Quantidade'] = df['Quantidade'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    df['Nr. nota'] = df['Nr. nota'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    if df['Prazo'].iloc[0] == "":
        df['Especificação do título'] = Utils.funcoes.sanitiza_especificacao_titulo(
        df['Especificação do título'])
    df['Tipo Mercado'] = Utils.funcoes.sanitiza_especificacao_titulo(df['Tipo Mercado'])
    df['Prazo'] = Utils.funcoes.sanitiza_especificacao_titulo(df['Prazo'])
    df['Prazo'].fillna("", inplace=True)
    df['Obs. (*)'] = Utils.funcoes.sanitiza_observacao(df['Obs. (*)'])

    tipotikect = ""
    try:
        if 'Unnamed: 0' in df.columns:
            tipotikect = df['Unnamed: 0']
            df['Unnamed: 0'] = df['Unnamed: 0'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    except:
        pass
    try:
        if 'Unnamed: 1' in df.columns:
            df['Unnamed: 1'] = df['Unnamed: 1'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    except:
        pass
    try:
        if 'Unnamed: 2' in df.columns:
            df['Unnamed: 2'] = df['Unnamed: 2'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    except:
        pass
    # ==============================================================================================
    # Coleta de dados por área de informação - Extraindo os dados das taxas e impostos
    # 51.691,400.509,72.516,568.597  - nota e data do pregão:
    # 450.341,32.576,639.253,544.276 - Resumo dos negócios, Resumo financeiro e Custos operacionais:
    # ==============================================================================================
    data = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False, stream=True,
    multiple_tables=True, pages='all', encoding="utf-8",
    area=(
    (51.691,400.509,72.516,568.597),
    (496.453,24.172,671.978,555.209))
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
    if 'Unnamed: 2' in df_gastos.columns:
        df_gastos['Unnamed: 2'] = df_gastos['Unnamed: 2'].apply(
        Utils.funcoes.sanitiza_moeda).astype('float')
        df_gastos['Unnamed: 2'].fillna(0, inplace=True)
    lista = list(
    df_gastos[df_gastos['Resumo dos Negócios'].str.contains("Valor das operações",na=False)].index)
    note_taxa = []

    # Obtem o número da conta na corretora
    conta = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False, stream=True,
    multiple_tables=True, pages='1', encoding="utf-8",
    area=(
    158.791,422.822,
    179.616,518.022)
    )
    conta = pd.concat(conta,axis=0,ignore_index=True)
    conta = conta['Código cliente'].iloc[0].strip().lstrip('0')
    conta = conta.split(" ")[1].strip().lstrip('0')

    # Verifica se a Nota de Corretagem já foi processada anteriormente
    #cpf = str(df['C.P.F./C.N.P.J/C.V.M./C.O.B.'][1])
    #if df['Tipo Mercado'].iloc[2].split(" ")[0] == "OPCAO":
    #    nome = conta + '_' + df_gastos['Data pregão'][0][6:10]
    #    nome += '_' + df_gastos['Data pregão'][0][3:5] + '_Opcoes.xlsx'
    #else:
    #    nome = conta + '_' + df_gastos['Data pregão'][0][6:10]
    #    nome += '_' + df_gastos['Data pregão'][0][3:5] + '_AVista.xlsx'
    #current_path = './Resultado/'
    #folder_prefix = str(
    #df['C.P.F./C.N.P.J/C.V.M./C.O.B.'][1] +'/'+ corretora +'/'+ df_gastos['Data pregão'][0][6:10])
    #folder_path = join(current_path, folder_prefix)
    #if exists(folder_path+'/'+nome):
    #    log.append(Utils.funcoes.verifica_nota_corretagem(folder_path,nome,item))
    #    Utils.funcoes.log_processamento(current_path,cpf,log)

    for current_row in lista:
        nota = df_gastos['Nr. nota'].iloc[current_row-8]
        data = datetime.strptime(df_gastos['Data pregão'].iloc[current_row-8], '%d/%m/%Y').date()
        total = df_gastos['Unnamed: 0'].iloc[current_row]
        vendas = df_gastos[
        'Unnamed: 0'].iloc[current_row-6] + df_gastos['Unnamed: 0'].iloc[current_row-3]
        liquidacao = df_gastos['Unnamed: 2'].iloc[current_row-5]
        registro = df_gastos['Unnamed: 2'].iloc[current_row-4]
        emolumentos = df_gastos['Unnamed: 2'].iloc[current_row+1]
        corretagem = df_gastos['Unnamed: 2'].iloc[current_row+4]
        imposto = df_gastos['Unnamed: 2'].iloc[current_row+7]
        irrf = df_gastos['Unnamed: 2'].iloc[current_row+8]
        outros = df_gastos['Unnamed: 2'].iloc[current_row+9]
        ir_daytrade = str(df_gastos['Unnamed: 2'].iloc[current_row+9])
        ir_daytrade = float(ir_daytrade.replace('.','').replace(',','.'))
        basecalculo = str(df_gastos['Unnamed: 1'].iloc[current_row+8])
        basecalculo = float(basecalculo.replace('.','').replace(',','.'))
        row_data = [
        nota,data,total,vendas,liquidacao,registro,emolumentos,corretagem,imposto,outros,
        emolumentos+liquidacao+registro,corretagem+imposto+outros,irrf,ir_daytrade,basecalculo]
        note_taxa.append(row_data)
    cols = [
    'Nota','Data','Total','Vendas','Liquidação','Registro','Emolumentos','Corretagem','Imposto',
    'Outros','Custos_Fin','Custos_Op','IRRF','IR_DT','BaseCalculo']
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
                #if df['Tipo Mercado'].iloc[current_row].split(" ")[0] == "OPCAO":
                #    nome = conta + '_' + data[6:10] + '_' + data[3:5] + '_Opcoes.xlsx'
                #else:
                #    nome = conta + '_' + data[6:10] + '_' + data[3:5] + '_AVista.xlsx'
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
        if df['Tipo Mercado'].iloc[current_row] == "VISTA":
            mercado = "VISTA"
        elif df['Tipo Mercado'].iloc[current_row] == "OPCAO DE COMPRA":
            mercado = "CALL"
        elif df['Tipo Mercado'].iloc[current_row] == "OPCAO DE VENDA":
            mercado = "PUT"
        elif df['Tipo Mercado'].iloc[current_row] == "EXERC OPC VENDA":
            mercado = "EXERC PUT"
        elif df['Tipo Mercado'].iloc[current_row] == "EXERC OPC COMPRA":
            mercado = "EXERC CALL"
        else:
            mercado = df['Tipo Mercado'].iloc[current_row]

        # Prazo de Vencimento da Opção
        if df['Prazo'].iloc[current_row] != "":
            monthrange = calendar.monthrange(
            2000 + int(df['Prazo'].iloc[current_row][3:]),int(df['Prazo'].iloc[current_row][0:2]))
            prazo = str(
            monthrange[1]) +'/' + str(int(df['Prazo'].iloc[current_row][0:2])) + '/' +str(
            2000 + int(df['Prazo'].iloc[current_row][3:]))
            prazo = datetime.strptime(prazo, '%d/%m/%Y').date()
        elif df['Prazo'].iloc[current_row] != "" and df[
        'Tipo Mercado'].iloc[current_row].lower() == "termo":
            prazo = int(df['Prazo'].iloc[current_row])
        else:
            prazo = ""

        # Exercicio de opção de compra/venda - atualizado em 15/09/2024
        if df['Tipo Mercado'].iloc[current_row].split(" ")[0] == "EXERC":
            exercicio = df['Especificação do título'].iloc[current_row]
            exercicio = exercicio.split(" ")[0][:-1]
        elif df['Tipo Mercado'].iloc[current_row].split(" ")[0] == "OPCAO":
            exercicio = df['Especificação do título'].iloc[current_row].split(" ")[0][0:4]
            try:
                if isinstance(tipotikect.iloc[current_row], str):
                    letras = ""
                    letras = tipotikect.iloc[current_row].split(" ")[0]
                    codigo  =  codigos[letras]
                    exercicio = exercicio + codigo
                elif len(df['Especificação do título'].iloc[current_row].split(" "))==2:
                    letras = ""
                    letras = df['Especificação do título'].iloc[current_row].split(" ")[1]
                    codigo  =  codigos[letras]
                    exercicio = exercicio + codigo
                else:
                    exercicio = Utils.funcoes.converte_opcao_ticket(
                    df['Especificação do título'].iloc[current_row])
            except ValueError:
                codigo = '11'
        else:
            exercicio = ""

        # Altera o número de dias de um contrato a Termo para a
        # data de vencimento desse contrato
        if mercado in "TERMOTermoTERMO":
            data_termo = datetime.datetime.strptime(data, "%m/%d/%y")
            prazo = data_termo + datetime.timedelta(days=prazo)
            prazo = datetime.strptime(prazo, '%d/%m/%Y').date()

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
        custo_financeiro,irrf_operacao = Utils.funcoes.custos_por_operacao(taxas_df,numero_nota,
        c_v,valor_total,operacao)
        irrf_operacao = irrf_operacao * 10

        # Susbstitui o nome do papel no pregão pelo seu respectivo código na B3 no padrão "XXXX3"
        # Caso seja uma opção de compra/venda o código continuará o mesmo
        if df['Tipo Mercado'].iloc[current_row].lower() in ["vista", "fracionario", "termo"]:
            stock_title = stock_title.split(" ")[0]
        elif df['Tipo Mercado'].iloc[current_row].split(" ")[0] == "EXERC" and df[
        'Especificação do título'].iloc[current_row].split(" ")[0][-1:] == "E":
            # if data > datetime.strptime("01/06/2024", '%d/%m/%Y').date():
            stock_title = df['Especificação do título'].iloc[current_row].split(" ")[0][0:4]
            try:
                if isinstance(tipotikect.iloc[current_row], str):
                    letras = ""
                    letras = tipotikect.iloc[current_row].split(" ")[0]
                    codigo  =  codigos[letras]
                    stock_title = stock_title + codigo
                elif len(df['Especificação do título'].iloc[current_row].split(" "))==2:
                    letras = ""
                    letras = df['Especificação do título'].iloc[current_row].split(" ")[1]
                    codigo  =  codigos[letras]
                    stock_title = stock_title + codigo
                else:
                    stock_title,log_nome_pregao = Utils.funcoes.nome_pregao_opcoes(
                    opcoes, stock_title, data)
                    if log_nome_pregao != temp:
                        temp = log_nome_pregao
                        log.append(log_nome_pregao)
            except ValueError:
                codigo = '11'
        elif df['Tipo Mercado'].iloc[current_row].split(" ")[0] == "EXERC" and df[
        'Especificação do título'].iloc[current_row].split(" ")[0][-1:] != "E":
            stock_title = df['Especificação do título'].iloc[current_row].split(" ")[0][0:4]
            string = df['Especificação do título'].iloc[current_row]
            letras = string[string.rfind('E')+1:]
            codigo  =  codigos[letras]
            stock_title = stock_title + codigo
            exercicio = string[:string.rfind('E')]
        elif df['Tipo Mercado'].iloc[current_row].split(" ")[0] == "OPCAO":
            stock_title = stock_title.split(" ")[0]
        else:
            stock_title,log_nome_pregao = Utils.funcoes.nome_pregao(acoes, stock_title, data)
            if log_nome_pregao != temp:
                temp = log_nome_pregao
                log.append(log_nome_pregao)

        #Calculando o preço médio de cada operação
        pm = Utils.funcoes.preco_medio(c_v,valor_total,custo_financeiro,quantidade)

        row_data = [
        corretora, cpf, numero_nota, data, c_v, stock_title, operacao, preco_unitario, quantidade,
        valor_total,custo_financeiro, pm, irrf_operacao,mercado,prazo,exercicio
        ]
        note_data.append(row_data)
    cols = [
    'Corretora','CPF', 'Nota', 'Data', 'C/V', 'Papel', 'Operacao','Preço', 'Quantidade', 'Total',
    'Custos_Fin','PM', 'IRRF','Mercado','Prazo','Exercicio'
    ]
    note_df = pd.DataFrame(data=note_data, columns=cols)

    #Agrupar os dados de preço e quantidade por ativo C/V em cada NC
    grouped = Utils.funcoes.agrupar(note_df)
    grouped = grouped[cols]

    # Seleção de papel isento de IR (IRRF e IRPF).
    note_data,log_isecao = Utils.funcoes.isencao_imposto_renda(taxas_df,grouped,note_data)
    log.append(log_isecao)
    cols = [
    'Corretora','CPF', 'Nota', 'Data', 'C/V', 'Papel', 'Operacao','Preço', 'Quantidade', 'Total',
    'Custos_Fin', 'PM', 'IRRF','Mercado','Prazo','Exercicio']
    note_df = pd.DataFrame(data=note_data, columns=cols)

    # Refaz o agrupamento para atualizar os dados de preço e quantidade
    # com a correção de compra/venda a maior no DayTrade
    grouped = Utils.funcoes.agrupar_btg(note_df)
    grouped = grouped[cols]

    # Agrupa as operações por tipo de operação (Normal ou Daytrade)
    try:
        normal_df,daytrade_df,result = Utils.funcoes.agrupar_operacoes(grouped,cols)
        # Insere o valor do IR para as operações de Daytrade"
        note_data,taxas_df,log_daytrade_ir = Utils.funcoes.daytrade_ir(result,taxas_df,
        note_data,grouped)
        log.append(log_daytrade_ir)
        cols = [
        'Corretora','CPF', 'Nota', 'Data', 'C/V', 'Papel', 'Operacao','Preço', 'Quantidade',
        'Total','Custos_Fin', 'PM', 'IRRF','Mercado','Prazo','Exercicio']
        note_df = pd.DataFrame(data=note_data, columns=cols)
        # Refaz o agrupamento para atualizar os dados de preço e quantidade
        grouped = Utils.funcoes.agrupar(note_df)
        grouped = grouped[cols]
    except ValueError:
        Utils.funcoes.agrupar_operacoes(grouped,cols)

    # Acrescentando os custos operacionais (Corretagem, Imposto e Outros)
    grouped = Utils.funcoes.custos_operacionais(grouped,taxas_df)

    # Obtendo o valor correto do preço unitário de cada operação
    grouped['Preço'] = grouped['Total'] / grouped['Quantidade']

    # Obtendo o valor correto do preço médio de cada operação
    grouped['PM'] = Utils.funcoes.preco_medio_correcao(grouped)

    # Inseri o número da conta na corretora
    grouped.insert(2,"Conta",int(conta),True)
    taxas_df.insert(0,"Conta",int(conta),True)

    # Agrupa as operações por tipo de trade
    # com correção de compra/venda a maior no DayTrade
    normal_df,daytrade_df = Utils.funcoes.agrupar_operacoes_correcao(grouped,cols)
    cols = [
    'Corretora','CPF','Conta','Nota', 'Data', 'C/V', 'Papel','Mercado','Preço', 'Quantidade',
    'Total', 'Custos_Fin', 'PM', 'IRRF','Prazo','Exercicio'
    ]
    if not normal_df.empty:
        normal_df = normal_df[cols]
    if not daytrade_df.empty:
        daytrade_df = daytrade_df[cols]

    # Cria o caminho completo de pastas/subpasta para salvar o resultado do processamento
    #current_path = './Resultado/'
    #folder_prefix = cpf+'/'+corretora+'/'+ano
    #folder_path = join(current_path, folder_prefix)
    log_move_resultado = Utils.funcoes.move_resultado(cpf)
    log.append(log_move_resultado)

    # Disponibiliza os dados coletados em um arquivo .xlsx separado por mês
    #Utils.funcoes.arquivo_separado(folder_path,nome,note_df,normal_df,daytrade_df,taxas_df)

    # Disponibiliza todos os dados coletados de todos os arquivos processados em um único arquivo
    # Utils.funcoes.arquivo_unico(current_path,cpf,note_df,normal_df,daytrade_df,taxas_df)
    Utils.funcoes.arquivo_unico(current_path,cpf,normal_df,daytrade_df)

    # Disponibiliza todos os dados coletados de todos os arquivos processados em um único arquivo
    if pagebmf != 0 and control != 0:
        Utils.Corretoras.btg_bmf.btg_bmf(corretora,filename,item,log,pagebmf,control)

    # Cria o caminho completo de pastas/subpastas para mover os arquivos já processados.
    log_move_saida = Utils.funcoes.move_saida(cpf,corretora,ano,item)
    log.append(log_move_saida)

    # Cria um arquivo de LOG para armazenar os dados do processamento
    log.append(
    'As Notas de Corretagem do arquivo "'+basename(item)+'" foram processadas com sucesso.\n')
    Utils.funcoes.log_processamento(current_path,cpf,log)
