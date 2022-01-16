#!/usr/bin/env python3
from openpyxl import load_workbook
import pandas as pd


print("Docker is running!")

file = r'Vendas-Comb.xlsx'

planilha = load_workbook(file)
sheet = planilha['Plan1']

'''
As the information is in pivot tables, to avoid having to open excel and having to assemble
a new tab with the information of all the years for each type of Sale, if it is necessary to create
a function to automatically retrieve data from predefined dynamic tables.
'''

def cache_to_df(cache):
    total = []
    cols = {}
    for i in cache.cacheFields:
        cols[i.name] = i.name

    for linhas in cache.records.r:
        linha = []
        for dat in linhas._fields:
            try:
                linha.append(dat.v)
            except AttributeError:
                linha.append(None)
        total.append(linha)
    df = pd.DataFrame(columns=cols, data=total)
    return df

# Confirm that the cache is for the desired tables.
contagem=-1
for table_din in sheet._pivots:
    contagem = contagem + 1
    if table_din.name == 'Tabela dinâmica1':
        num_tab_dinam1 = contagem
        print('Encontrada')
    elif table_din.name == 'Tabela dinâmica3':
        num_tab_dinam3 = contagem
        print('Encontrada')
    else:
        print('Tabelas dinâmicas não encontradas')


#Getting info of the pivot table
pivot_oil = sheet._pivots[num_tab_dinam1].cache
pivot_diesel = sheet._pivots[num_tab_dinam3].cache

#Raw dataframe
df_diesel = cache_to_df(pivot_diesel)
df_oil = cache_to_df(pivot_oil)

#Starting the "Transform" step:
    
df_diesel = df_diesel.astype({"COMBUSTÍVEL": str, "ANO": str, "REGIÃO": str, "ESTADO": str})
df_oil = df_oil.astype({"COMBUSTÍVEL": str, "ANO": str, "REGIÃO": str, "ESTADO": str})

change_ano_o = {'0': '2000', '1': '2001','2': '2002',
                              '3': '2003', '4': '2004', '5': '2005',
                              '6': '2006', '7': '2007',
                              '8': '2008', '9': '2009', '10': '2010',
                              '11': '2011', '12': '2012',
                              '13': '2013', '14': '2014', '15': '2015',
                              '16': '2016', '17': '2017', '18': '2018',
                              '19': '2019', '20': '2020'}

change_ano_d = {'0': '2013', '1': '2014','2': '2015',
                              '3': '2016', '4': '2017', '5': '2018',
                              '6': '2019', '7': '2020'}


change_comb_o = {'0': 'ETANOL HIDRATADO', '1': 'GASOLINA C','2': 'GASOLINA DE AVIAÇÃO',
                              '3': 'GLP', '4': 'ÓLEO COMBUSTÍVEL', '5': 'ÓLEO DIESEL',
                              '6': 'QUEROSENE DE AVIAÇÃO', '7': 'QUEROSENE ILUMINANTE'}

change_comb_d = {'0': 'ÓLEO DIESEL (OUTROS)', '1': 'ÓLEO DIESEL MARÍTIMO','2': 'ÓLEO DIESEL S-10',
                              '3': 'ÓLEO DIESEL S-1800', '4': 'ÓLEO DIESEL S-500'}

#Putting the real year for every correspondece number
df_oil['ANO'] = df_oil['ANO'].map(change_ano_o)
df_diesel['ANO'] = df_diesel['ANO'].map(change_ano_d)

#Putting the real fuel for every correspondece number
df_oil['COMBUSTÍVEL'] = df_oil['COMBUSTÍVEL'].map(change_comb_o)
df_diesel['COMBUSTÍVEL'] = df_diesel['COMBUSTÍVEL'].map(change_comb_d)

#Getting rid of nan type
df_oil = df_oil.fillna(0)
df_diesel = df_diesel.fillna(0)

#Rounding the 'TOTAL' to compare after the transformations
df_oil['TOTAL'] = df_oil['TOTAL'].round(2)
df_diesel['TOTAL'] = df_diesel['TOTAL'].round(2)



def transform(df_fuel):
    
   df_fuel.rename(columns = {'Jan':'January', 'Fev':'February',
                                    'Mar':'March', 'Abr':'April',
                                    'Mai':'May', 'Jun':'June',
                                    'Jul':'July', 'Ago':'August',
                                    'Set':'September', 'Out':'October',
                                    'Nov':'November', 'Dez':'December'}, inplace=True)
    
   abbrev_estad = {'0': 'AC', '1': 'AL','2': 'AP',
                                  '3': 'AM', '4': 'BA', '5': 'CE',
                                  '6': 'DF', '7': 'ES',
                                  '8': 'GO', '9': 'MA', '10': 'MT',
                                  '11': 'MS', '12': 'MG',
                                  '13': 'PA', '14': 'PB', '15': 'PR',
                                  '16': 'PE', '17': 'PI', '18': 'RJ',
                                  '19': 'RN', '20': 'RS',
                                  '21': 'RO', '22': 'RR', '23': 'SC',
                                  '24': 'SP', '25': 'SE', '26': 'TO'}
        
   df_fuel['UF'] = df_fuel['ESTADO'].map(abbrev_estad)
   
   df_fuel = df_fuel.drop('TOTAL', axis = 1)
   df_fuel = df_fuel.drop('REGIÃO', axis = 1)
   df_fuel = df_fuel.drop('ESTADO', axis = 1)
   
   df_fuel = df_fuel.melt(id_vars=['COMBUSTÍVEL', 'ANO', 'UNIDADE','UF'],
                     var_name = 'MONTH', value_name = 'VOLUME');
   
   
   df_fuel['created_at'] = pd.to_datetime(df_fuel['ANO'].astype(str) + df_fuel['MONTH'].astype(str), format='%Y%B')
    
   df_fuel['year_month'] = df_fuel['created_at'].dt.strftime('%Y-%m')
        
   df_fuel.VOLUME = df_fuel['VOLUME'].round(2);
   
   df_fuel = df_fuel.rename(columns = {'UNIDADE':'unit','VOLUME':'volume','COMBUSTÍVEL':'product','UF':'uf'});

   df_fuel = df_fuel[['year_month', 'uf', 'product', 'unit', 'volume', 'created_at']];
   

   return(df_fuel)


df_oil_final = transform(df_oil)
df_diesel_final = transform(df_diesel)


Total_ini_o = df_oil['TOTAL'].sum()
Total_final_o = df_oil_final['volume'].sum()


Total_ini_d = df_diesel['TOTAL'].sum()
Total_final_d = df_diesel_final['volume'].sum()


#Comparacao do somatorio de valores antes e depois do ETL
if (Total_ini_o/Total_final_o) >= 0.999999 and (Total_ini_d/Total_final_d) >= 0.999999 :
    print('Checkpoint de somatório correto!')
    print('Oil Table:' )
    print(df_oil_final)
    print('Diesel Table:' )
    print(df_diesel_final)
else:
    print('Valores totais divergindo')
    

