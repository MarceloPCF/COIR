# Importações padrão de bibliotecas Python
from datetime import datetime
from os.path import join, basename, exists

# Importações de terceiros
import tabula
import pandas as pd

# Importação das funções definidas em outro módulo
import Utils.funcoes

# ==================================================================================================
# Processamento de notas de corretagens BM&F das corretoras do grupo XP (XP, Rico e Clear)
# Rotina para extração de dados no novo layout das notas de corretagens - início em 01/2024
# ==================================================================================================
def nao_validada_bmf(corretora,filename,item,log,page,control):
    # ==============================================================================================
    # Coleta de dados por área de informação - Extraindo os dados das operações na B3
    # 51.777,445.138,68.913,556.888   - Nota e data do pregão
    # 126.278,444.393,145.648,558.377 - CPF
    # 181.408,31.663,629.153,562.848  - Informações de compra e venda
    # ==============================================================================================
    data = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False, stream=True,
    multiple_tables=True, pages=page, encoding="utf-8",
    area=(
    (51.777,445.138,68.913,556.888),
    (126.278,444.393,145.648,558.377),
    (181.408,31.663,629.153,562.848))
    )
    df = pd.concat(data,axis=0,ignore_index=True)
    df['Nr. nota'] = df['Nr. nota'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    df['C/V'] = Utils.funcoes.sanitiza_especificacao_titulo(df['C/V'])
    df['Mercadoria'] = df['Mercadoria'].str.replace(' ','',regex=False)
    df['Quantidade'] = df['Quantidade'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    df['Preço/Ajuste'] = df['Preço/Ajuste'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    df['Tipo Negócio'] = Utils.funcoes.sanitiza_especificacao_titulo(df['Tipo Negócio'])
    df['Vlr de Operação/Ajuste'] = df[
    'Vlr de Operação/Ajuste'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    df['D/C'] = Utils.funcoes.sanitiza_especificacao_titulo(df['D/C'])
    df['Taxa Operacional'] = df[
    'Taxa Operacional'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    if 'Unnamed: 0' in df.columns: #CPF do investidor
        df['Unnamed: 0'] = Utils.funcoes.sanitiza_especificacao_titulo(df['Unnamed: 0'])
    if 'Unnamed: 1' in df.columns:
        df['Unnamed: 1'] = df['Unnamed: 1'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    if 'Unnamed: 2' in df.columns:
        df['Unnamed: 2'] = df['Unnamed: 2'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    # ==============================================================================================
    # Coleta de dados por área de informação - Extraindo os dados das taxas e impostos
    # 46.484,442.159,68.797,561.90   - Nota e data do pregão
    # 636.603,32.408,718.553,561.358 - Resumo dos negócios, Resumo financeiro e Custos operacionais
    # ==============================================================================================
    data = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False, stream=True,
    multiple_tables=True, pages=page, encoding="utf-8",
    area=(
    (51.033,445.138,70.403,559.123),
    (636.603,32.408,718.553,561.358))
    )
    df_gastos = pd.concat(data,axis=0,ignore_index=True)
    #print(df_gastos.to_string())

    df_gastos['Nr. nota'] = df_gastos[
    'Nr. nota'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
    if 'Unnamed: 0' in df_gastos.columns:
        df_gastos['Unnamed: 0'] = df_gastos['Unnamed: 0'].apply(Utils.funcoes.sanitiza_nota_bmf)
        df_gastos['Unnamed: 0'] = df_gastos[
        'Unnamed: 0'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        df_gastos['Unnamed: 0'].fillna(0, inplace=True)
    if 'Unnamed: 1' in df_gastos.columns:
        df_gastos['Unnamed: 1'].fillna(0, inplace=True)
    try:
        if 'Unnamed: 2' in df_gastos.columns:
            df_gastos['Unnamed: 2'] = df_gastos['Unnamed: 2'].apply(Utils.funcoes.sanitiza_nota_bmf)
            df_gastos['Unnamed: 2'] = df_gastos[
            'Unnamed: 2'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
            df_gastos['Unnamed: 2'].fillna(0, inplace=True)
    except ValueError:
        if 'Unnamed: 2' in df_gastos.columns:
            df_gastos['Unnamed: 2'].fillna(0, inplace=True)
    if 'Unnamed: 3' in df_gastos.columns:
        df_gastos['Unnamed: 3'] = df_gastos[
        'Unnamed: 3'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        df_gastos['Unnamed: 3'].fillna(0, inplace=True)
    if 'Unnamed: 4' in df_gastos.columns:
        df_gastos['Unnamed: 4'] = df_gastos['Unnamed: 4'].apply(Utils.funcoes.sanitiza_nota_bmf)
        df_gastos['Unnamed: 4'] = df_gastos[
        'Unnamed: 4'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        df_gastos['Unnamed: 4'].fillna(0, inplace=True)
    if 'Unnamed: 5' in df_gastos.columns:
        df_gastos['Unnamed: 5'] = df_gastos['Unnamed: 5'].apply(Utils.funcoes.sanitiza_nota_bmf)
        df_gastos['Unnamed: 5'] = df_gastos[
        'Unnamed: 5'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        df_gastos['Unnamed: 5'].fillna(0, inplace=True)
    if 'Unnamed: 6' in df_gastos.columns:
        df_gastos['Unnamed: 6'] = df_gastos['Unnamed: 6'].apply(Utils.funcoes.sanitiza_nota_bmf)
        df_gastos['Unnamed: 6'] = df_gastos[
        'Unnamed: 6'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        df_gastos['Unnamed: 6'].fillna(0, inplace=True)
    if 'Unnamed: 7' in df_gastos.columns:
        df_gastos['Unnamed: 7'] = df_gastos[
        'Unnamed: 7'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        df_gastos['Unnamed: 7'].fillna(0, inplace=True)
    if 'Unnamed: 8' in df_gastos.columns:
        df_gastos['Unnamed: 8'] = df_gastos[
        'Unnamed: 8'].apply(Utils.funcoes.sanitiza_moeda).astype('float')
        df_gastos['Unnamed: 8'].fillna(0, inplace=True)
    lista = list(df_gastos[df_gastos['Venda disponível'].str.contains("IRRF",na=False)].index)
    note_taxa = []

    # Obtem o número da conta na corretora 144.903,446.628,162.783,559.123
    try:
        conta = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False,
        stream=True, multiple_tables=True, pages='1', encoding="utf-8",
        area=(156.078,425.768,177.683,521.127))
        conta = pd.concat(conta,axis=0,ignore_index=True)
        conta = conta['Unnamed: 0'].iloc[0].strip().lstrip('0')
    except KeyError:
        conta = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False,
        stream=True, multiple_tables=True, pages='1', encoding="utf-8",
        area=(144.903,445.138,162.783,556.143))
        conta = pd.concat(conta,axis=0,ignore_index=True)
        conta = conta['Unnamed: 0'].iloc[0].strip().lstrip('0')
    except ValueError:
        conta = tabula.read_pdf(filename, pandas_options={'dtype': str}, guess=False,
        stream=True, multiple_tables=True, pages='1', encoding="utf-8",
        area=(144.903,445.138,162.783,556.143))
        conta = pd.concat(conta,axis=0,ignore_index=True)
        conta = conta['Unnamed: 0'].iloc[0].strip().lstrip('0')

    # Verifica se a Nota de Corretagem já foi processada anteriormente
    current_path = './Resultado/'

    

    for current_row in lista:
        nota = df_gastos['Nr. nota'].iloc[current_row-2]
        data = datetime.strptime(
        df_gastos['Data pregão'].iloc[current_row-2], '%d/%m/%Y').date()
        irrf = str(df_gastos['Unnamed: 2'].iloc[current_row+1])
        if irrf != "nan":
            #irrf = irrf.split("|")[0]
            irrf = irrf.split("|", maxsplit=1)[0]
            irrf = float(irrf.replace('.','').replace(',','.'))
        else:
            irrf = "0"
        venda_disponivel = df_gastos['Unnamed: 2'].iloc[current_row-1]
        compra_disponivel = df_gastos['Unnamed: 4'].iloc[current_row-1]
        #venda_opcoes = df_gastos['Unnamed: 5'].iloc[current_row-1]
        #compra_opcoes = df_gastos['Unnamed: 6'].iloc[current_row-1]
        #valor_negocios = df_gastos['Unnamed: 7'].iloc[current_row-1]
        ir_daytrade = df_gastos['Unnamed: 4'].iloc[current_row+1]
        corretagem = df_gastos['Unnamed: 5'].iloc[current_row+1]
        taxa_registro = df_gastos['Unnamed: 6'].iloc[current_row+1]
        emolumentos = df_gastos['Unnamed: 7'].iloc[current_row+1]
        outros_custos = df_gastos['Compra disponível'].iloc[current_row+3]
        outros_custos = float(outros_custos.replace('.','').replace(',','.'))
        imposto = df_gastos['Unnamed: 4'].iloc[current_row+3]
        #ajuste_posicao = df_gastos['Unnamed: 5'].iloc[current_row+3]
        #ajuste_daytrade = df_gastos['Unnamed: 6'].iloc[current_row+3]
        #total_custos_operacionais = df_gastos['Unnamed: 7'].iloc[current_row+3]
        outros = df_gastos['Unnamed: 0'].iloc[current_row+5]
        #ir_operacional = df_gastos['Unnamed: 2'].iloc[current_row+5]
        #total_conta_investimento = df_gastos['Unnamed: 4'].iloc[current_row+5]
        #total_conta_normal = df_gastos['Unnamed: 5'].iloc[current_row+5]
        #total_liquido = df_gastos['Unnamed: 6'].iloc[current_row+5]
        #total_liquido_nota = df_gastos['Unnamed: 7'].iloc[current_row+5]
        liquidacao = 0
        basecalculo = 0
        row_data = [nota,data,compra_disponivel,venda_disponivel,liquidacao,taxa_registro,
        emolumentos,corretagem,imposto,outros,liquidacao+imposto+outros,
        corretagem+imposto+outros,irrf,ir_daytrade,basecalculo]
        note_taxa.append(row_data)
    cols = ['Nota','Data','Total','Vendas','Liquidação','Registro','Emolumentos','Corretagem'
    ,'Imposto','Outros','Custos_Fin','Custos_Op','IRRF','IR_DT','BaseCalculo']
    taxas_df = pd.DataFrame(data=note_taxa, columns=cols)
    indexnames = taxas_df[((taxas_df['Custos_Fin'] == 0) & (taxas_df['Custos_Op'] == 0))].index
    taxas_df.drop(indexnames ,inplace=True)
    taxas_df = taxas_df.drop_duplicates(subset=['Nota','Data'], keep='last', ignore_index=True)

    cont_notas = len(taxas_df['Nota'])
    if cont_notas > 1:
        log.append('Serão processadas ' + str(cont_notas) + ' notas de índice Futuros ou BMF.\n')
    else:
        log.append('Será processada ' + str(cont_notas) + ' notas de índice de Futuros ou BMF.\n')

    # Incluir aqui a etapa para obter lista de linhas de cada operação
    operacoes = list(df[df['C/V'].isin(['C','V'])].index)
    # vendas = list(df[df['C/V'].isin(['V'])].index)

    if len(operacoes) == 0 and control == 1:
        log.append(
        'Nota(s) de Corretagem(ns) apenas com ajustes de posição, não foi contabilizada.\n')
        cpf = df['Unnamed: 0'].iloc[current_row-1]
        Utils.funcoes.log_processamento(current_path,cpf,log)
        return
    if len(operacoes) == 0 and control == 2:
        log.append(
        'Nota(s) de Corretagem(ns) apenas com ajustes de posição, não foi contabilizada.\n')
        current_path = './Resultado/'
        #cpf = df['Unnamed: 0'].iloc[current_row-1]
        #data = df['Data pregão'].iloc[current_row-2]
        ano = data[6:10]
        #nome = ''
        #folder_prefix = cpf+'/'+corretora+'/'+ano
        #folder_path = join(current_path, folder_prefix)
        log_move_saida = Utils.funcoes.move_saida(cpf,corretora,ano,item)
        log.append(log_move_saida)
        log_move_resultado = Utils.funcoes.move_resultado(cpf)
        log.append(log_move_resultado)
        Utils.funcoes.log_processamento(current_path,cpf,log)
        return

    note_data = []
    numero_nota = 0
    cpf = ''
    #nome = ''
    ano = ''
    #temp = ''
    for current_row in operacoes:
        cell_value = df['Nr. nota'].iloc[current_row-2]
        if cell_value > 0:
            numero_nota = df['Nr. nota'].iloc[current_row-2]
            data = df['Data pregão'].iloc[current_row-2]
            if ano == '':
                cpf = df['Unnamed: 0'].iloc[current_row-1]
                #nome = conta + '_' + data[6:10] + '_' + data[3:5] + '.xlsx'
                ano = data[6:10]
            data = datetime.strptime(df['Data pregão'].iloc[current_row-2], '%d/%m/%Y').date()

        if df['Tipo Negócio'].iloc[current_row] in 'NORMALDAY TRADE':

            # Tipo de operação (Compra ou Venda)
            c_v = df['C/V'].iloc[current_row].strip()

            # Nome do ativo no pregão
            mercadoria = df['Mercadoria'].iloc[current_row].strip()
            #tipo_negocio = df['Tipo Negócio'].iloc[current_row]
            operacao = df['Tipo Negócio'].iloc[current_row]
            if operacao == "DAY TRADE":
                operacao = "DayTrade"
            else:
                operacao = "Normal"

            # Preço unitário da operação de cada mercadoria por nota de corretagem
            preco_unitario = df['Preço/Ajuste'].iloc[current_row]

            # Quantidade operada de cada mercadoria por nota de corretagem
            quantidade = df['Quantidade'].iloc[current_row]

            # Valor total de cada operação por nota de corretagem
            valor_total,id,mult = Utils.funcoes.mercadoria_ticket(
            mercadoria,preco_unitario,quantidade)

            # Valor de corretagem por cada mercadoria operada
            corretagem = df['Taxa Operacional'].iloc[current_row]

            # Alterao nome da variável para manter a compatibilidade com o script de ações
            stock_title = mercadoria

            # Obtem o valor individualizado da taxa de Registro e Emolumentos de cada operação
            datalimite = datetime.strptime("01/08/2021", '%d/%m/%Y').date()
            if data >= datalimite:
                registro_emol,mercado = Utils.funcoes.taxas_registro_emol(
                operacao,df['Mercadoria'].iloc[current_row][:3],stock_title)
            else:
                registro_emol,mercado = Utils.funcoes.taxas_registro_emol_old(
                operacao,df['Mercadoria'].iloc[current_row][:3],stock_title)

            # Dividindo os custos
            taxas_df_custos_fin = None
            taxas_df_corretagem = None
            custo_financeiro = 0
            if corretagem == 0:
                for i in taxas_df.index:
                    if numero_nota == taxas_df['Nota'].iloc[i] and data == taxas_df['Data'].iloc[i]:
                        custo_financeiro = (
                        taxas_df['Custos_Fin'].iloc[i] + taxas_df['Registro'].iloc[i] + taxas_df[
                        'Emolumentos'].iloc[i])/len(operacoes)
                        break
            else:
                for i in taxas_df.index:
                    if taxas_df['Corretagem'].iloc[i] == 0:
                        taxas_df_corretagem = 1
                    else:
                        taxas_df_corretagem = taxas_df['Corretagem'].iloc[i]
                    if taxas_df['Custos_Fin'].iloc[i] == 0:
                        taxas_df_custos_fin = 1
                    else:
                        taxas_df_custos_fin = taxas_df['Custos_Fin'].iloc[i]
                    if numero_nota == taxas_df['Nota'].iloc[i] and data == taxas_df[
                    'Data'].iloc[i] and df['Mercadoria'].iloc[
                    current_row][:3] in "WINwinINDindWDOwdoDOLdol":
                        custo_financeiro = (corretagem/taxas_df_corretagem) * (
                        taxas_df_custos_fin + taxas_df['Registro'].iloc[i] + taxas_df[
                        'Emolumentos'].iloc[i])
                        break
                    if numero_nota==taxas_df['Nota'].iloc[i] and data==taxas_df['Data'].iloc[i]:
                        custo_financeiro = (registro_emol*quantidade) + (
                        (corretagem/taxas_df_corretagem)*taxas_df_custos_fin)
                        break

            irrf_operacao = 0
            ir_daytrade = 0

            # Calculando o preço médio de cada operação - Para operações de Futuros não se caucula PM
            pm = 0

            row_data = [corretora, cpf, numero_nota, data, c_v, stock_title,operacao,preco_unitario,
            quantidade, valor_total, custo_financeiro + corretagem,pm,irrf_operacao,ir_daytrade,id,
            mult,mercado]
            note_data.append(row_data)

    cols = ['Corretora','CPF', 'Nota', 'Data', 'C/V', 'Papel', 'Operacao','Preço','Quantidade',
    'Total','Custos_Fin','PM','IRRF','IR_DT','ID','FATOR','Mercado']
    note_df = pd.DataFrame(data=note_data, columns=cols)

    # Contabiliza a quantidade de vendas nas operações DayTrade e Normal
    note_df = Utils.funcoes.ir_bmf(cont_notas,note_df,taxas_df,row_data,note_data)

    # Agrupar os dados de preço e quantidade por cada ativo C/V em cada nota de corretagem
    grouped = Utils.funcoes.agrupar_bmf(note_df)

    # Inseri o número da conta na corretora
    grouped.insert(2,"Conta",int(conta),True)
    taxas_df.insert(0,"Conta",int(conta),True)

    # Agrupa as operações por tipo de trade com correção de compra/venda a maior no DayTrade
    cols = ['Corretora','CPF','Conta','Nota','Data','C/V','Papel','Mercado','Preço','Quantidade',
    'Total','Custos_Fin','PM','IRRF']
    normal_df,daytrade_df = Utils.funcoes.agrupar_operacoes_correcao(grouped,cols)
    if normal_df.empty is False:
        normal_df = normal_df[cols]
    if daytrade_df.empty is False:
        daytrade_df = daytrade_df[cols]

    # Cria o caminho completo de pastas/subpasta para salvar o resultado do processamento
    #current_path = './Resultado/'
    #folder_prefix = cpf+'/'+corretora+'/'+ano
    #folder_path = join(current_path, folder_prefix)
    if control == 2:
        log_move_resultado = Utils.funcoes.move_resultado(cpf)
        log.append(log_move_resultado)

    # Cria o caminho completo de pastas/subpasta para salvar o resultado do processamento
#    current_path = './Resultado/'
#    folder_prefix = cpf+'/'+corretora+'/'+ano
#    folder_path = join(current_path, folder_prefix)
#    if control == 2:
#        log_move_resultado,pagebmf = Utils.funcoes.move_resultado(
#        folder_path,cpf,nome,item,pagebmf=1)
#        log.append(log_move_resultado)

    # Não exportar os dados de 'ID','FATOR', sem utilidade no momento
    cols = ['Corretora','CPF','Nota','Data','C/V','Papel','Operacao','Preço','Quantidade','Total',
    'Custos_Fin','PM','IRRF','Mercado']
    note_df = note_df[cols]

    # Disponibiliza os dados coletados em um arquivo .xlsx separado por mês
#    if control == 2:
#        Utils.funcoes.arquivo_separado(folder_path,nome,note_df,normal_df,daytrade_df,taxas_df)
#    else:
#        Utils.funcoes.arquivo_separado_bmf(folder_path,nome,note_df,normal_df,daytrade_df,taxas_df)

    # Disponibiliza todos os dados coletados de todos os arquivos processados em um único arquivo
    # Utils.funcoes.arquivo_unico(current_path,cpf,note_df,normal_df,daytrade_df,taxas_df)
    current_path = './Resultado/'
    Utils.funcoes.arquivo_unico(current_path,cpf,normal_df,daytrade_df)

    # Cria o caminho completo de pastas/subpastas para mover os arquivos já processados.
    if control == 2:
        log_move_saida = Utils.funcoes.move_saida(cpf,corretora,ano,item)
        log.append(log_move_saida)

    # Cria um arquivo de LOG para armazenar os dados do processamento
    if control == 2:
        log.append(
        'As Notas de Corretagens do arquivo "'+basename(item)+'" foram processadas com sucesso.\n')
        Utils.funcoes.log_processamento(current_path,cpf,log)
