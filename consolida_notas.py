
'''
Partes da nota de corretagem
df[1]
    cabecalho
df[1]
    negocios realizados
df[2]
    coluna 0
        resumo dos negocios
        especificacoes diversas
    coluna 1:
        resumo financeiro
            CBLC
            Bovespa/Soma
            Corretagem/Despesas
'''

import sys
import tabula
from os.path import join
import glob
import pandas as pd

def comma_str_to_float(comma_str):
    return float(comma_str.strip().replace('.', '').replace(',', '.'))

###############################################################################
def get_header(nota):
    
    df_head = nota[0]
    df_head = df_head.iloc[1]
    num_nota = df_head[2].split(' ')[0]
    dt_pregao = df_head[3]
    return num_nota, dt_pregao

###############################################################################
def get_negocios(nota):
    
    df_neg = nota[1]
    if df_neg.shape[1] == 12:
        df_neg[5] = df_neg[5] + ' ' + df_neg[6].fillna('')
        df_neg.drop(6, axis=1, inplace=True)
        
    colunas = ['Q', 'Negociação', 'C/V', 'Tipo Mercado', 'Prazo',
               'Especificação do título', 'Obs', 'Quantidade', 'Preço/Ajuste',
               'Valor/Ajuste', 'D/C']
    
    df_neg.columns = colunas
    df_neg = df_neg.iloc[1:]
    df_neg = df_neg[['C/V', 'Tipo Mercado', 'Especificação do título',
                     'Quantidade', 'Preço/Ajuste', 'Valor/Ajuste', 'D/C']]
    df_neg['Quantidade'] = df_neg['Quantidade'].apply(int)
    df_neg['Preço/Ajuste'] = df_neg['Preço/Ajuste'].apply(comma_str_to_float)
    df_neg['Valor/Ajuste'] = df_neg['Valor/Ajuste'].apply(comma_str_to_float)
    return df_neg

###############################################################################
def get_custos(nota):
    
    df_custos = nota[2]
    df_custos = df_custos.iloc[:, 1:3]
    df_custos.columns = ['item', 'valor']
    
    liquidacao = comma_str_to_float(df_custos.loc[df_custos['item'] == 'Taxa de liquidação'].iloc[0, 1])
    emolumentos = comma_str_to_float(df_custos.loc[df_custos['item'] == 'Emolumentos'].iloc[0, 1])
    corretagem = comma_str_to_float(df_custos.loc[df_custos['item'] == 'Corretagem'].iloc[0, 1])
    iss = comma_str_to_float(df_custos.loc[df_custos['item'] == 'ISS(SÃO PAULO)'].iloc[0, 1])
    outros_bovespa = comma_str_to_float(df_custos.loc[df_custos['item'] == 'Outras'].iloc[0, 1])

    
    return liquidacao, emolumentos, corretagem, iss, outros_bovespa
    
###############################################################################
def lista_arquivos(diretorio_de_notas):
    return glob.glob(join(dir_notas, '*.pdf'))

###############################################################################
def pdf_to_df(arq_nota):
    print('Processando:', arq_nota)
    sys.stdout.flush()
    nota_corr = tabula.read_pdf(arq_nota, multiple_tables=True)
    
    negocios_part = get_negocios(nota_corr)
    qtd_oper = len(negocios_part)
    
    num_nota, dt_pregao = get_header(nota_corr)
    negocios_part['NUM_NOTA'] = num_nota
    negocios_part['DT_PREG'] = dt_pregao
    
    liquidacao, emolumentos, corretagem, iss, outros_bovespa = get_custos(nota_corr)
    negocios_part['liquidacao'] = liquidacao/qtd_oper
    negocios_part['emolumentos'] = emolumentos/qtd_oper
    negocios_part['corretagem'] = corretagem/qtd_oper
    negocios_part['iss'] = iss/qtd_oper
    negocios_part['outros_bovespa'] = outros_bovespa/qtd_oper
    
    return negocios_part

###############################################################################
def consolida_notas(lista_notas):
    df_notas = pd.concat(lista_notas)
    df_notas.set_index('DT_PREG', inplace=True)
    df_notas.index = pd.to_datetime(df_notas.index, dayfirst=True)
    df_notas.sort_index(inplace=True)
    
    return df_notas

###############################################################################
def calcula_preco_medio(row):
#    print(row)
    custos = ['liquidacao', 'emolumentos', 'corretagem',
              'iss', 'outros_bovespa']
    custo = row[custos]

    if row['C/V'] == 'C':
        valor_ajustado = row['Valor/Ajuste'] + custo
    elif row['C/V'] == 'V':
        valor_ajustado = row['Valor/Ajuste'] - custo
    
    return valor_ajustado/row['Quantidade']
    
###############################################################################
dir_notas = r'O:\OneDrive\Documents\Contas\itaucorretora\notas_corretagem'

notas = map(pdf_to_df, lista_arquivos(dir_notas))
notas = consolida_notas(notas)
#notas['PRC_MED'] = notas.apply(calcula_preco_medio, axis=1)

notas.to_excel(join(dir_notas, 'resumo.xlsx'))
