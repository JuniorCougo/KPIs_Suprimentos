# ACOMPANHAMENTO DE PEDIDOS DE FORNECEDORES
import pandas as pd
import datetime

#DIAS EM ATRASO
# Carregar os arquivos Excel
df1 = pd.read_excel("C:\\Users\\sar8577\\Documents\\AA_DATA FRAME_PYTHON\\df_pedidos_de_compras_II_MXM.xlsx")
df2 = pd.read_excel("C:\\Users\\sar8577\\Documents\\AA_DATA FRAME_PYTHON\\base_grupos_de_cotacao_consolidado.xlsx")

print(f"Pedidos de Compras Possui {df1.shape[0]} linhas e {df1.shape[1]} colunas.")
print(f"Base Grupo de Cotação {df2.shape[0]} linhas e {df2.shape[1]} colunas.")


# Filtrar df1 e renomear as colunas para simplificar a manipulação
df1 = df1[df1['Unnamed: 36'].isin(["Aprovado", "Atend. Parcial", "Atend. Total"]) & 
          df1['Unnamed: 39'].isin(["Sem Classificação", "Dispensa", "Inexigível"])]
df1 = df1.rename(columns={
    'Unnamed: 0': 'Filial',
    'Unnamed: 6': 'Pedido',
    'Unnamed: 7': 'Setor',
    'Unnamed: 11': 'Comprador',
    'Unnamed: 16': 'Condição de Pagamento',
    'Unnamed: 18': 'Fornecedor',
    'Unnamed: 19': 'Data do Pedido',
    'Unnamed: 20': 'Linha Pedido',
    'Unnamed: 21': 'Código_MXM',
    'Unnamed: 22': 'Descrição',
    'Unnamed: 23': 'Unidade de Medida',
    'Unnamed: 25': 'Qtd. Solicitada',
    'Unnamed: 26': 'Qtd Atendida',
    'Unnamed: 27': 'Preço Unitário',
    'Unnamed: 29': 'Num.Req.Compra',
    'Unnamed: 33': 'Data de Entrega',
    'Unnamed: 34': 'Tipo de Requisição',
    'Unnamed: 35': 'Grupo de Cotação',
    'Unnamed: 36': 'Status do Pedido',
    'Unnamed: 37': 'Status do Item',
    'Unnamed: 38': 'Saldo do Item',
    'Unnamed: 39': 'Classificação do Mapa',
    'Unnamed: 41': 'Vlr. Total do Pedido',
})

# Conversão da coluna 'Data do Pedido' para datetime e criação de colunas de mês e ano
df1['Data do Pedido'] = pd.to_datetime(df1['Data do Pedido'], errors='coerce')
df1['Mes Compra'] = df1['Data do Pedido'].dt.month
df1['Ano Compra'] = df1['Data do Pedido'].dt.year

# Remoção das colunas 'Unnamed' não renomeadas
df1 = df1.filter(regex='^(?!Unnamed:)')

# Criar as colunas 'STATUS DA ENTREGA' e 'DIAS EM ATRASO' com a regra adicional
data_atual = datetime.datetime.now().date()
df1['Data de Entrega'] = pd.to_datetime(df1['Data de Entrega']).dt.date
df1['STATUS DA ENTREGA'] = df1.apply(
    lambda x: 'Entrega Concluída' if x['Status do Pedido'] == 'Atend. Total'
    else ('Entrega em Atraso' if x['Data de Entrega'] < data_atual else 'Entrega no Prazo'),
    axis=1
)
df1['DIAS EM ATRASO'] = df1.apply(
    lambda x: '' if x['STATUS DA ENTREGA'] == 'Entrega Concluída'
    else (data_atual - x['Data de Entrega']).days if x['STATUS DA ENTREGA'] == 'Entrega em Atraso' else 0,
    axis=1
)


# Ajustes para colunas 'Status do Item' e 'Classificação do Mapa'
df1['Status do Item'] = df1['Status do Item'].fillna('Item Pendente de Entrega Total')
df1['Classificação do Mapa'] = df1['Classificação do Mapa'].replace('Sem Classificação', 'Ata/Contrato')
df1['Status do Pedido'] = df1['Status do Pedido'].replace('Aprovado', 'Pedido Pendente de Entrega Total')

# Remover duplicados com base na chave concatenada
df1['Chave_Concatenada'] = df1['Pedido'].astype(str) + "_" + df1['Linha Pedido'].astype(str) + "_" + df1['Código_MXM'].astype(str) + "_" + df1['Qtd. Solicitada'].astype(str)
df1 = df1.drop_duplicates(subset=['Chave_Concatenada'])

# Criar colunas de valores e % de atendimento por item
df1['% DE AT. POR ITEM'] = (df1['Qtd Atendida'] / df1['Qtd. Solicitada']) * 100
df1['Valor Solicitado Item'] = df1['Qtd. Solicitada'] * df1['Preço Unitário']
df1['Valor Atendido Item'] = df1['Qtd Atendida'] * df1['Preço Unitário']
df1['Valor Pendente de Atendimento Item'] = df1['Saldo do Item'] * df1['Preço Unitário']

# Merge de df1 com df2 para obter o GRUPO DE COTACAO CONSOLIDADO
df2_reduzido = df2[['GRUPO DE COTACAO', 'GRUPO CONSOLIDADO']].drop_duplicates()
df1 = pd.merge(df1, df2_reduzido, left_on='Grupo de Cotação', right_on='GRUPO DE COTACAO', how='left').drop(columns=['GRUPO DE COTACAO'])
df1 = df1.rename(columns={'GRUPO CONSOLIDADO': 'GRUPO DE COTACAO CONSOLIDADO'})


print("Sheet_Pedidos Detalhados OK!")

# Agrupar as informações para a planilha consolidada
df_consolidado = df1.groupby(
    [
        'Filial', 'Pedido', 'Data do Pedido', 'Fornecedor', 'Classificação do Mapa', 
        'Status do Pedido', 'Data de Entrega', 'STATUS DA ENTREGA', 'DIAS EM ATRASO'
    ],
    as_index=False
).agg({
    'GRUPO DE COTACAO CONSOLIDADO': 'first',
    'Valor Solicitado Item': 'sum',
    'Valor Atendido Item': 'sum',
    'Valor Pendente de Atendimento Item': 'sum'
})
df_consolidado['% de ATENDIMENTO'] = (df_consolidado['Valor Atendido Item'] / df_consolidado['Valor Solicitado Item']) * 100

# Garantir que a coluna DIAS EM ATRASO seja do tipo numérico
df_consolidado['DIAS EM ATRASO'] = pd.to_numeric(df_consolidado['DIAS EM ATRASO'], errors='coerce')
# Ordenar o DataFrame consolidado pela coluna 'DIAS EM ATRASO' do maior para o menor
df_consolidado = df_consolidado.sort_values(by='DIAS EM ATRASO', ascending=False)

print("Sheet_Pedidos Consolidados OK!")


# Criar a tabela de frequência de compras por pedido
df = df1.copy()
df['Mes'] = df['Data do Pedido'].dt.month
df['Ano'] = df['Data do Pedido'].dt.year
df = df.drop_duplicates(subset=['Pedido', 'Código_MXM'])
df['Tipo do Item'] = df['Código_MXM'].apply(
    lambda x: 'Consumo' if str(x).startswith('C') else
              'Estoque' if str(x).startswith('E') else
              'Patrimônio' if str(x).startswith('PT') else
              'Serviço' if str(x).startswith(('PS', 'OB')) else
              'Desconhecido'
)
frequencia_compras = df.groupby(['Filial', 'Código_MXM', 'Mes', 'Ano']).agg({
    'Código_MXM': 'count',
    'Descrição': 'first',
    'Unidade de Medida': 'first',
    'Preço Unitário': ['min', 'max', 'mean'],
    'Tipo do Item': 'first'
}).reset_index()
frequencia_compras.columns = ['Filial', 'Código_MXM', 'Mes', 'Ano', 'Frequencia', 'Descrição', 'Unidade de Medida', 'Valor_Compra_Min', 'Valor_Compra_Max', 'Valor_Compra_Medio', 'Tipo do Item']
frequencia_compras['Variacao_%'] = (frequencia_compras['Valor_Compra_Max'] - frequencia_compras['Valor_Compra_Min']) / frequencia_compras['Valor_Compra_Min'] * 100
frequencia_compras = frequencia_compras[['Filial', 'Tipo do Item', 'Código_MXM', 'Descrição', 'Unidade de Medida', 'Mes', 'Ano', 'Frequencia', 'Valor_Compra_Min', 'Valor_Compra_Max', 'Variacao_%', 'Valor_Compra_Medio']]
frequencia_compras = frequencia_compras.sort_values(by='Frequencia', ascending=False)
#frequencia_compras = frequencia_compras.query('(Mes in [9, 10]) and (Ano in [2024, 2025])')

print("Sheet_Frequencia de Compras OK!")


# Exportar para o arquivo Excel
formato_data_hora = "%Y%m%d_%H%M%S"
data_hora_atual = datetime.datetime.now().strftime(formato_data_hora)
nome_arquivo = f"Pedidos_De_Compras_v4_{data_hora_atual}.xlsx"
caminho_saida = "C:/Users/sar8577/Downloads/"

with pd.ExcelWriter(caminho_saida + nome_arquivo, engine='openpyxl') as writer:
    df1.to_excel(writer, sheet_name='Pedidos Detalhados', index=False)
    df_consolidado.to_excel(writer, sheet_name='Consolidado', index=False)
    frequencia_compras.to_excel(writer, sheet_name='Frequencia Compras', index=False)



print("Processo Concluído Com Sucesso!")