#Frequencia de compras por pedidos de comrpas x variação de custo do pedido (entendo que deve ser por requisição porque uma requisição pode gerar mais de um pedido)

import pandas as pd
import datetime

# Carregando o arquivo Excel
df1 = pd.read_excel("C:\\Users\\sar8577\\Documents\\AA_DATA FRAME_PYTHON\\df_pedidos_de_compras_II_MXM.xlsx")

# Renomear colunas
df = df1.rename(columns={
    'Unnamed: 0': 'Filial',
    'Unnamed: 5': 'Pedido',
    'Unnamed: 19': 'Data do Pedido', # Data_Compra
    'Unnamed: 21': 'Código_MXM', #Item
    'Unnamed: 22': 'Descrição', #ok
    'Unnamed: 23': 'Unidade de Medida',
    #'Unnamed: 26': 'Valor de Compra'  # Renomeando a coluna de valor de compra
    'Unnamed: 27': 'Preço Unitário',
    'Unnamed: 29': 'Num.Req.Compra',
})


# Data_Compra = Data do Pedido / Item = Código_MXM' / Valor de Compra = Preço Unitário


# Converter a coluna 'Data_Compra' para o tipo datetime
df['Data do Pedido'] = pd.to_datetime(df['Data do Pedido'], format='%d/%m/%Y', errors='coerce')

# Criar colunas separadas para mês e ano
df['Mes'] = df['Data do Pedido'].dt.month
df['Ano'] = df['Data do Pedido'].dt.year

# Criar a chave concatenando as colunas, pedido e item /// 'Unnamed: 6' é número do pedido, porém entendi que devo alterar para número da requisição alterou pra 5 não sei porque 
df['Chave'] = df['Pedido'].astype(str) + '_' + df['Código_MXM'].astype(str)
#'Unnamed: 28' = Número da requisição

# Remover linhas duplicadas com base na chave, mantendo apenas a primeira ocorrência
df = df.drop_duplicates(subset=['Chave'], keep='first')

# Remover a coluna 'Chave' após a remoção de duplicatas
df = df.drop('Chave', axis=1)

# Função para definir o 'Tipo do Item'
def tipo_item(codigo_item):
    # Converter para string e tratar valores nulos
    codigo_item = str(codigo_item) if pd.notnull(codigo_item) else ''
    
    if codigo_item.startswith('C'):
        return 'Consumo'
    elif codigo_item.startswith('E'):
        return 'Estoque'
    elif codigo_item.startswith('PT'):
        return 'Patrimônio'
    elif codigo_item.startswith('PS'):
        return 'Serviço'
    elif codigo_item.startswith('OB'):
        return 'Serviço'
    else:
        return 'Desconhecido'

# Aplicar a função para a nova coluna
df['Tipo do Item'] = df['Código_MXM'].apply(tipo_item)

# Agrupar por Filial, Item, Mês e Ano e calcular os valores necessários
frequencia_compras = df.groupby(['Filial', 'Código_MXM', 'Mes', 'Ano']) \
    .agg({
        'Código_MXM': 'count',  # Contagem de compras
        'Descrição': 'first',  # Primeira ocorrência da descrição
        'Unidade de Medida': 'first',  # Primeira ocorrência da unidade de medida
        'Preço Unitário': ['min', 'max', 'mean'],  # Mínimo, máximo e média de valor de compra
        'Tipo do Item': 'first'  # Primeiro valor da coluna 'Tipo do Item'
    })

# Ajustar os nomes das colunas para remover a hierarquia
frequencia_compras.columns = ['Frequencia', 'Descrição', 'Unidade de Medida', 'Valor_Compra_Min', 'Valor_Compra_Max', 'Valor_Compra_Medio', 'Tipo do Item']

# Resetar o índice para que 'Filial', 'Item', 'Mes' e 'Ano' voltem a ser colunas normais
frequencia_compras = frequencia_compras.reset_index()


# ... (seu código anterior, sem alterações significativas)

def calcular_variacao(minimo, maximo):
    if minimo == 0:
        return "Não aplicável"  # Ou np.nan para valores ausentes
    else:
        return ((maximo - minimo) / minimo) * 100

# ... (seu código anterior até a parte do cálculo da variação)

frequencia_compras['Variacao_%'] = frequencia_compras.apply(
    lambda x: calcular_variacao(x['Valor_Compra_Min'], x['Valor_Compra_Max']), axis=1
)

# ... (restante do seu código)




##############################


# Reordenar as colunas do DataFrame
frequencia_compras = frequencia_compras[['Filial', 'Tipo do Item', 'Código_MXM', 'Descrição', 'Unidade de Medida',  'Mes', 'Ano', 
                                         'Frequencia', 'Valor_Compra_Min', 'Valor_Compra_Max', 'Variacao_%',  
                                         'Valor_Compra_Medio']]

# Renomear colunas
frequencia_compras = frequencia_compras.rename(columns={
    'Filial': 'FILIAL',
    'Tipo do Item': 'TIPO',
    'Código_MXM': 'Código_MXM',
    'Descrição': 'DESCRIÇÃO',
    'Unidade de Medida': 'UN',
    'Mes': 'MÊS',
    'Ano': 'ANO',
    'Frequencia':'FREQUÊNCIA DE COMPRAS POR PC',
    'Valor_Compra_Min':'MÍN. (R$)',
    'Valor_Compra_Max':'MÁX. (R$)',
    'Variacao_%': 'VARIAÇÃO (%)',
    'Valor_Compra_Medio': ' CUSTO MÉDIO MÊS (R$)'
})

# Ordenar o DataFrame com base na coluna 'FREQUÊNCIA DE COMPRAS POR RC' do maior para o menor
frequencia_compras = frequencia_compras.sort_values(by='FREQUÊNCIA DE COMPRAS POR PC', ascending=False)

# Exportar o DataFrame para o arquivo Excel
formato_data_hora = "%Y%m%d_%H%M%S"
data_hora_atual = datetime.datetime.now().strftime(formato_data_hora)
nome_arquivo = f"Frequencia_Compras_por_PC_X_Variação_de_Custo{data_hora_atual}.xlsx"
caminho_saida = "C:/Users/sar8577/Downloads/"
frequencia_compras.to_excel(caminho_saida + nome_arquivo, index=False)

# Exibir o resultado
print(frequencia_compras)
