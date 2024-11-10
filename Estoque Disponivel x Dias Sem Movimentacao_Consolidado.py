# Estoque Filial x Estoque Disponível x Dias Sem Movimentação
import pandas as pd
import openpyxl
import os
from datetime import datetime
print('Executando................Estoque Disponível x Dias sem Movimentação')
# Caminhos para os arquivos
caminho_df1 = r"C:\Users\sar8577\Documents\AA_DATA FRAME_PYTHON\df_estoque_disponivel_mxm.xlsx"
caminho_df2 = r"C:\Users\sar8577\Documents\AA_DATA FRAME_PYTHON\df_dias_sem_movimentaca_mxm.xlsx"

# Lendo os DataFrames
df1 = pd.read_excel(caminho_df1)
df2 = pd.read_excel(caminho_df2)

# Obtendo as datas de modificação dos arquivos
data_modificacao_df1 = os.path.getmtime(caminho_df1)
data_modificacao_df2 = os.path.getmtime(caminho_df2)

# Formatando as datas para visualização
data_modificacao_df1 = datetime.fromtimestamp(data_modificacao_df1).strftime('%Y-%m-%d %H:%M:%S')
data_modificacao_df2 = datetime.fromtimestamp(data_modificacao_df2).strftime('%Y-%m-%d %H:%M:%S')

print(f"Estoque disponível possui {df1.shape[0]} linhas e {df1.shape[1]} colunas. Última modificação: {data_modificacao_df1}")
print(f"Dias sem Movimentação possui {df2.shape[0]} linhas e {df2.shape[1]} colunas. Última modificação: {data_modificacao_df2}")

# Excluindo as primeiras duas linhas do df2
df2 = df2.iloc[2:]

# Verificando o tipo das colunas para garantir que as chaves estejam corretas
print(df1['EST_CODIGO'].dtype)
print(df1['ITEM'].dtype)
print(df2['Unnamed: 0'].dtype)
print(df2['Unnamed: 8'].dtype)

########################### Fase 2 ###########################

####Tive que remover as duplicatas por causa os itens estão duplicando no relátório disponnível
# Criando a chave primária com as colunas FILIAL, EST_DESCRICAO e ITEM
df1['chave_primaria'] = df1['FILIAL'].astype(str) + df1['EST_DESCRICAO'].astype(str) + df1['ITEM'].astype(str)
# Remover duplicatas com base na chave primária
df1 = df1.drop_duplicates(subset='chave_primaria', keep='first')
# Agora a chave está pronta sem duplicatas/####Tive que remover as duplicatas por causa os itens estão duplicando no relátório disponnível

########

# Criando a chave primária em cada DataFrame (similar ao CONCATENAR do Excel)
df1['chave_primaria'] = df1['FILIAL'].astype(str) + df1['ITEM'].astype(str)
df2['chave_primaria'] = df2['Unnamed: 0'].astype(str) + df2['Unnamed: 8'].astype(str)

# Criando a coluna "CUSTO FINAL" no df1 com base nas colunas 'Unnamed: 20' e 'Unnamed: 22' do df2
df2['Unnamed: 20'] = pd.to_numeric(df2['Unnamed: 20'], errors='coerce')
df2['Unnamed: 22'] = pd.to_numeric(df2['Unnamed: 22'], errors='coerce')
df1['CUSTO FINAL'] = df2['Unnamed: 20'] / df2['Unnamed: 22']

# Substituir os valores nulos e zero de "CUSTO FINAL" por "PRECOCUSTO" ou "PRECOCOMPRA"
df1['CUSTO FINAL'] = df1['CUSTO FINAL'].fillna(0)
df1['CUSTO FINAL'] = df1['CUSTO FINAL'].where(df1['CUSTO FINAL'] != 0, df1['PRECOCUSTO'])
df1['CUSTO FINAL'] = df1['CUSTO FINAL'].where(df1['PRECOCUSTO'] != 0, df1['PRECOCOMPRA'])
df1['CUSTO FINAL'] = df1['CUSTO FINAL'].fillna(df1['PRECOCOMPRA'])

# Mesclando os DataFrames (similar ao PROCV)
df_final = pd.merge(df1, df2, on='chave_primaria', how='left')

# Criando a coluna "VALOR TOTAL" no df_final
df_final['VALOR TOTAL'] = df_final['CUSTO FINAL'] * df_final['ESTOQUEFISICOREAL']

# Renomeando colunas
df_final = df_final.rename(columns={
    'Unnamed: 19': 'DIAS SEM MOV.',
    'Unnamed: 18': 'DT DA ULT. MOV.',
    'SALDOFISICOESTOQUE': 'ESTQ. + ESTQ TRÂNSITO CD'
})

# Excluindo as filiais 1, 9 e 25 da coluna 'FILIAL'
df_final = df_final[~df_final['FILIAL'].isin([9,  25,1, 65, 7,68, 71])] #~ invete a operação# 25,1

# Selecionando apenas as colunas desejadas
colunas_desejadas = [
    'FILIAL', 'EST_CODIGO', 'EST_DESCRICAO', 'CODIGOCOTACAO', 'DESCCOTACAO', 'DT DA ULT. MOV.', 'DIAS SEM MOV.',
    'ITEM', 'PRODUTO', 'UNIDADE', 'QTDSOLICITADA', 'ESTQ. + ESTQ TRÂNSITO CD', 'ESTOQUEFISICOREAL', 
    'ESTOQUETRANSITO', 'DEMANDAXSALDO', 'STATUSDISPONIBILIDADE', 'CUSTO FINAL', 'VALOR TOTAL'
]
#Retirar as colunas:'PRECOCUSTO', 'PRECOCOMPRA', para não confundir 
# Remover 'SALDOSUPPLY', 'SALDOPENDENTECOMPRADIRETA', 'SALDOPENDENTEPC', por causa da dos itens duplicados até resolver

df_final = df_final[colunas_desejadas]

df_final = df_final.sort_values(by='DIAS SEM MOV.', ascending=False).reset_index(drop=True)


# Exibir o DataFrame final
#print(df_final.head())

####FASE 3####

# Filtrar os valores de "DIAS SEM MOV." maiores ou iguais a 180
df_resumo = df_final[df_final['DIAS SEM MOV.'] >= 180]

# Criar uma tabela dinâmica (pivot table) com "FILIAL" e "EST_DESCRICAO" como rótulos, 
# somar "VALOR TOTAL" e contar o número de SKUs distintos
df_pivot = df_resumo.pivot_table(
    index=['FILIAL', 'EST_DESCRICAO'], 
    values=['VALOR TOTAL', 'PRODUTO'],  # Coluna 'PRODUTO' será contada
    aggfunc={'VALOR TOTAL': 'sum', 'PRODUTO': 'nunique'}  # 'nunique' conta valores únicos
).reset_index()

# Renomear a coluna de contagem para 'TOTAL DE SKU's'
df_pivot = df_pivot.rename(columns={'PRODUTO': "TOTAL DE SKU's"})

# Ordenar do maior para o menor com base em "VALOR TOTAL"
df_pivot = df_pivot.sort_values(by='VALOR TOTAL', ascending=False).reset_index(drop=True)

# Exibir o resultado
print('Dias Sem Movimentção')
print(df_pivot.head())


####################SOMA DO SALDO DISPONIVEL   ###########
# Filtrar apenas as linhas com "STATUSDISPONIBILIDADE" igual a "Estoque Disponivel"
df_saldo_disponivel = df_final[df_final['STATUSDISPONIBILIDADE'] == 'Estoque Disponivel']

# Criar a tabela dinâmica (pivot table) com "FILIAL" e "EST_DESCRICAO" como rótulos,
# contando o total de SKU's e somando o valor total
df_pivot_saldo = df_saldo_disponivel.pivot_table(
    index=['FILIAL', 'EST_DESCRICAO'], 
    values=['VALOR TOTAL', 'PRODUTO'],  # Coluna 'PRODUTO' será contada
    aggfunc={'VALOR TOTAL': 'sum', 'PRODUTO': 'nunique'}  # 'nunique' conta valores únicos
).reset_index()

# Renomear a coluna de contagem para 'TOTAL DE SKU's'
df_pivot_saldo = df_pivot_saldo.rename(columns={'PRODUTO': "TOTAL DE SKU's"})

# Ordenar do maior para o menor com base em "VALOR TOTAL"
df_pivot_saldo = df_pivot_saldo.sort_values(by='VALOR TOTAL', ascending=False).reset_index(drop=True)

# Exibir o resultado
print('Saldo Disponível')
print(df_pivot_saldo.head())


############### INICIAR EXPORTAÇÃO ###################

# Criar um formato de data e hora personalizado
formato_data_hora = "%Y%m%d_%H%M%S"
data_hora_atual = datetime.now().strftime(formato_data_hora)

# Construir o nome do arquivo com a data e hora
nome_arquivo = f"Estoque_Disponivel_x_Dias_Sem_Movimentacao_{data_hora_atual}.xlsx"

# Definir o caminho de saída
caminho_saida = "C:/Users/sar8577/Downloads/"

# Combinar o caminho e o nome do arquivo
caminho_completo_arquivo = caminho_saida + nome_arquivo

# Exportar os DataFrames para o arquivo Excel
with pd.ExcelWriter(caminho_completo_arquivo, engine='openpyxl') as writer:
    df_final.to_excel(writer, sheet_name='Estoque Disp. x Dias Mov.', index=False)
    df_pivot.to_excel(writer, sheet_name='Sum. Dias Sem Mov. 180 Dias', index=False)
    df_pivot_saldo.to_excel(writer, sheet_name='Sum.Saldo Disponivel', index=False)

# Exibir as primeiras linhas de df_final como confirmação
print('df_final(Detalhado)')
print(df_final.head())
print('Exportado Com Sucesso!!!')

