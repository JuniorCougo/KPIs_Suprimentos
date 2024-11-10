import pandas as pd
import openpyxl
import os
from datetime import datetime
print('Executando................Estoque Disponível x Dias sem Movimentação')

caminho_df1 = r"C:x"
caminho_df2 = r"xlsx"

df1 = pd.read_excel(caminho_df1)
df2 = pd.read_excel(caminho_df2)

data_modificacao_df1 = os.path.getmtime(caminho_df1)
data_modificacao_df2 = os.path.getmtime(caminho_df2)


data_modificacao_df1 = datetime.fromtimestamp(data_modificacao_df1).strftime('%Y-%m-%d %H:%M:%S')
data_modificacao_df2 = datetime.fromtimestamp(data_modificacao_df2).strftime('%Y-%m-%d %H:%M:%S')

print(f"Estoque disponível possui {df1.shape[0]} linhas e {df1.shape[1]} colunas. Última modificação: {data_modificacao_df1}")
print(f"Dias sem Movimentação possui {df2.shape[0]} linhas e {df2.shape[1]} colunas. Última modificação: {data_modificacao_df2}")


df2 = df2.iloc[2:]

print(df1['EST_CODIGO'].dtype)
print(df1['ITEM'].dtype)
print(df2['Unnamed: 0'].dtype)
print(df2['Unnamed: 8'].dtype)


df1['chave_primaria'] = df1['FILIAL'].astype(str) + df1['EST_DESCRICAO'].astype(str) + df1['ITEM'].astype(str)

df1 = df1.drop_duplicates(subset='chave_primaria', keep='first')



df1['chave_primaria'] = df1['FILIAL'].astype(str) + df1['ITEM'].astype(str)
df2['chave_primaria'] = df2['Unnamed: 0'].astype(str) + df2['Unnamed: 8'].astype(str)

df2['Unnamed: 20'] = pd.to_numeric(df2['Unnamed: 20'], errors='coerce')
df2['Unnamed: 22'] = pd.to_numeric(df2['Unnamed: 22'], errors='coerce')
df1['CUSTO FINAL'] = df2['Unnamed: 20'] / df2['Unnamed: 22']


df1['CUSTO FINAL'] = df1['CUSTO FINAL'].fillna(0)
df1['CUSTO FINAL'] = df1['CUSTO FINAL'].where(df1['CUSTO FINAL'] != 0, df1['PRECOCUSTO'])
df1['CUSTO FINAL'] = df1['CUSTO FINAL'].where(df1['PRECOCUSTO'] != 0, df1['PRECOCOMPRA'])
df1['CUSTO FINAL'] = df1['CUSTO FINAL'].fillna(df1['PRECOCOMPRA'])

df_final = pd.merge(df1, df2, on='chave_primaria', how='left')


df_final['VALOR TOTAL'] = df_final['CUSTO FINAL'] * df_final['ESTOQUEFISICOREAL']

df_final = df_final.rename(columns={
    'Unnamed: 19': 'DIAS SEM MOV.',
    'Unnamed: 18': 'DT DA ULT. MOV.',
    'SALDOFISICOESTOQUE': 'ESTQ. + ESTQ TRÂNSITO CD'
})


df_final = df_final[~df_final['FILIAL'].isin([9,  25,1, 65, 7,68, 71])] #~ invete a operação# 25,1

colunas_desejadas = [
    'FILIAL', 'EST_CODIGO', 'EST_DESCRICAO', 'CODIGOCOTACAO', 'DESCCOTACAO', 'DT DA ULT. MOV.', 'DIAS SEM MOV.',
    'ITEM', 'PRODUTO', 'UNIDADE', 'QTDSOLICITADA', 'ESTQ. + ESTQ TRÂNSITO CD', 'ESTOQUEFISICOREAL', 
    'ESTOQUETRANSITO', 'DEMANDAXSALDO', 'STATUSDISPONIBILIDADE', 'CUSTO FINAL', 'VALOR TOTAL'
]
df_final = df_final[colunas_desejadas]

df_final = df_final.sort_values(by='DIAS SEM MOV.', ascending=False).reset_index(drop=True)


####FASE 3####

df_resumo = df_final[df_final['DIAS SEM MOV.'] >= 180]

# Criar uma tabela dinâmica (pivot table) com "FILIAL" e "EST_DESCRICAO" como rótulos, 
# somar "VALOR TOTAL" e contar o número de SKUs distintos
df_pivot = df_resumo.pivot_table(
    index=['FILIAL', 'EST_DESCRICAO'], 
    values=['VALOR TOTAL', 'PRODUTO'],  # Coluna 'PRODUTO' será contada
    aggfunc={'VALOR TOTAL': 'sum', 'PRODUTO': 'nunique'}  # 'nunique' conta valores únicos
).reset_index()

df_pivot = df_pivot.rename(columns={'PRODUTO': "TOTAL DE SKU's"})

df_pivot = df_pivot.sort_values(by='VALOR TOTAL', ascending=False).reset_index(drop=True)

print('Dias Sem Movimentção')
print(df_pivot.head())

df_saldo_disponivel = df_final[df_final['STATUSDISPONIBILIDADE'] == 'Estoque Disponivel']

df_pivot_saldo = df_saldo_disponivel.pivot_table(
    index=['FILIAL', 'EST_DESCRICAO'], 
    values=['VALOR TOTAL', 'PRODUTO'],  # Coluna 'PRODUTO' será contada
    aggfunc={'VALOR TOTAL': 'sum', 'PRODUTO': 'nunique'}  # 'nunique' conta valores únicos
).reset_index()

df_pivot_saldo = df_pivot_saldo.rename(columns={'PRODUTO': "TOTAL DE SKU's"})

df_pivot_saldo = df_pivot_saldo.sort_values(by='VALOR TOTAL', ascending=False).reset_index(drop=True)

print('Saldo Disponível')
print(df_pivot_saldo.head())

formato_data_hora = "%Y%m%d_%H%M%S"
data_hora_atual = datetime.now().strftime(formato_data_hora)


nome_arquivo = f"Estoque_Disponivel_x_Dias_Sem_Movimentacao_{data_hora_atual}.xlsx"
caminho_saida = "C://Downloads/"
caminho_completo_arquivo = caminho_saida + nome_arquivo

with pd.ExcelWriter(caminho_completo_arquivo, engine='openpyxl') as writer:
    df_final.to_excel(writer, sheet_name='Estoque Disp. x Dias Mov.', index=False)
    df_pivot.to_excel(writer, sheet_name='Sum. Dias Sem Mov. 180 Dias', index=False)
    df_pivot_saldo.to_excel(writer, sheet_name='Sum.Saldo Disponivel', index=False)

print('df_final(Detalhado)')
print(df_final.head())
print('Exportado Com Sucesso!!!')

