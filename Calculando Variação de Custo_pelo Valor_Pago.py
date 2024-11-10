#df_Pedido de Compra x Valor Pago.xlsx"# Caetano Gerou este relatório

import pandas as pd
import datetime


import datetime# Lendo apenas a tabela '2_PEDIDOS'
#df = pd.read_excel("C:\\Users\\sar8577\\Desktop\\LMC's_2025\\CUSTO de ITENS_2025.xlsx", sheet_name='2_PEDIDOS')
df = pd.read_excel("C:\\Users\\sar8577\\Documents\\AA_DATA FRAME_PYTHON\\df_Pedido de Compra x Valor Pago.xlsx")
# Filtrando os dados para o ano de 2024 (assumindo que a coluna 'Dt Pedido' está no formato de data)
df_2024 = df[df['Dt Pedido'].dt.year == 2024]

# Agrupando os dados e calculando as estatísticas#Alterei a descrição para descricao sem ç~
resultados = df_2024.groupby(['Item', 'Descricao', 'Un']).agg(
    Valor_da_ultima_compra=('Preço Unitario', 'last'),
    Valor_Maximo=('Preço Unitario', 'max'),
    Valor_Minimo=('Preço Unitario', 'min'),
    Mediana=('Preço Unitario', 'median'),
    Media=('Preço Unitario', 'mean'),
    Item=('Item', 'first'),
    Descricao=('Descricao', 'first'),
    Un=('Un', 'first')
)

# Nova ordem das colunas
nova_ordem = ['Item', 'Descricao', 'Un','Valor_Minimo','Valor_Maximo', 'Variacao_%','Valor_da_ultima_compra', 'Mediana', 'Media']

# Reordenando as colunas
resultados = resultados.reindex(columns=nova_ordem)

# Calculando a variação percentual entre o valor mínimo e máximo
resultados['Variacao_%'] = ((resultados['Valor_Maximo'] - resultados['Valor_Minimo']) / resultados['Valor_Minimo']) * 100

######

# ... (seu código existente)
#este código é para formatar em % formatado mas não ficou bom
# Calculando a variação percentual entre o valor mínimo e máximo e formatando como porcentagem
#def formatar_porcentagem(valor):
    #return f"{valor:.2f}%"

#resultados['Variacao_%'] = ((resultados['Valor_Maximo'] - resultados['Valor_Minimo']) / resultados['Valor_Minimo']) * 100
#resultados['Variacao_%'] = resultados['Variacao_%'].apply(formatar_porcentagem)

######
# Exportando para o Excel
#####Exportar#

# Criar um formato de data e hora personalizado
formato_data_hora = "%Y%m%d_%H%M%S"
data_hora_atual = datetime.datetime.now().strftime(formato_data_hora)

# Construir o nome do arquivo com a data e hora e informações adicionais
nome_arquivo = f"Variação_Custo_2024_pelo_valor_pago{data_hora_atual}.xlsx"

# Caminho completo para o arquivo
caminho_saida = "C:/Users/sar8577/Downloads/"

# Exportar o DataFrame para o arquivo Excel
resultados.to_excel(caminho_saida + nome_arquivo, index=False)

print(df.head(10))
