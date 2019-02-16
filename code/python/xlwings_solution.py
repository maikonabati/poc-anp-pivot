#!/usr/bin/env python
# coding: utf-8

# In[57]:


import os
import urllib
import xlwings as xw
import pandas as pd
import psycopg2
import numpy as np


# In[58]:


# Funções para banco
def fnc_open_connection():
    con = psycopg2.connect(
        host='XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX', 
        database='XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX', 
        user='XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX', 
        password='XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
    )
    return con

def fnc_query_insert(sql,tipo):
    try:        
        con = fnc_open_connection()        
        cur = con.cursor()
        cur.execute(sql)        
        con.commit()
        if tipo=='com':
            return cur.fetchall()
        else:
            return cur.rowcount
    except RuntimeError as error:
        return error
    finally:        
        cur.close()
        con.close()  


# In[59]:


# Nome do arquivo
nome_arquivo = '../../dist/diesel.xls'


# In[60]:


# URL do arquivo
url = 'http://www.anp.gov.br/images/DADOS_ESTATISTICOS/Vendas_de_Combustiveis/Vendas_de_Combustiveis_m3.xls'
file_name, headers = urllib.request.urlretrieve(url)


# In[61]:


# Carrega arquivos
wb_vba = xw.Book('../../dist/VBA.xlsm')
wb = xw.Book(file_name)

# Roda VBA de apoio
macro = wb_vba.macro('click')
macro()

# Remove arquivo antigo
try:
    os.remove(nome_arquivo)
except:
    print("O arquivo não existe!")
    
# Salva resultado
wb.save(nome_arquivo)

# Fecha o Excel
app = xw.apps.active 
app.quit()


# In[70]:


# Carrega dataframe
df_diesel_v2 = pd.DataFrame(columns=['COMBUSTÍVEL', 'ANO', 'REGIÃO', 'ESTADO', 'UNIDADE', 'Jan', 'Fev','Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'TOTAL'])
for num in range(0,6):
    df_diesel_v1 = pd.read_excel('../../dist/diesel.xls',sheet_name=num)    
    df_diesel_v2 = df_diesel_v2.append(df_diesel_v1, ignore_index=True)

# Pivot
pivot = pd.pivot_table(
    df_diesel_v2, values=['Jan', 'Fev','Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],    
    columns=['COMBUSTÍVEL', 'ANO', 'REGIÃO', 'ESTADO', 'UNIDADE']
)

df_diesel_v2 = pd.DataFrame(pivot)
df_diesel_v2.insert(loc=0, column='unidade', value=df_diesel_v2.index.get_level_values(5))
df_diesel_v2.insert(loc=0, column='estado', value=df_diesel_v2.index.get_level_values(4))
df_diesel_v2.insert(loc=0, column='regiao', value=df_diesel_v2.index.get_level_values(3))
df_diesel_v2.insert(loc=0, column='ano', value=df_diesel_v2.index.get_level_values(2))
df_diesel_v2.insert(loc=0, column='produto', value=df_diesel_v2.index.get_level_values(1))
df_diesel_v2.insert(loc=0, column='mes', value=df_diesel_v2.index.get_level_values(0))
df_diesel_v2.columns = ['mes', 'produto', 'ano', 'regiao', 'estado', 'unidade', 'volume']
df_diesel_v2.reset_index(drop=True, inplace=True)


# In[72]:


# Conversáo do mês
de_para_mes = {'Jan':1, 'Fev':2,'Mar':3, 'Abr':4, 'Mai':5, 'Jun':6, 'Jul':7, 'Ago':8, 'Set':9, 'Out':10, 'Nov':11, 'Dez':12}

mes_convertido = []
for mes in df_diesel_v2['mes']:
    mes_convertido.append(de_para_mes.get(mes))

df_diesel_v2['mes_int'] = np.array(mes_convertido)


# In[73]:


# Insere no PostgreSQL
sql_insert = 'truncate table etl.tab_vendas_combustiveis; insert into etl.tab_vendas_combustiveis(ano,mes,estado,produto,unidade,vol_demanda_m3,ano_mes_foto) values'

for index,row in df_diesel_v2.iterrows():
    sql_insert += str(
        "("+str(row['ano'])+
        ","+str(row['mes_int'])+
        ",'"+str(row['estado'])+
        "','"+str(row['produto'])+
        "','"+str(row['unidade'])+
        "',"+str(row['volume'])+
        ",'"+str(row['ano'])+'_'+str(row['mes'])+"'),")

retorno = str(fnc_query_insert(sql_insert[:-1],'sem'))
print('Total de registros inseridos: '+str(retorno))
print('\nFim!')

