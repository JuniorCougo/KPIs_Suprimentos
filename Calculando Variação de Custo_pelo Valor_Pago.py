#df_Pedido de Compra x Valor Pago.xlsx"

import pandas as pd
import datetime


import datetime
#df = pd.read("C:')
df = pd.l("C")

df_2024 = df[df['Dt Pedido'].dt.year == 2024]

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

nova_ordem = ['Item', 'Descricao', 'Un','Valor_Minimo','Valor_Maximo', 'Variacao_%','Valor_da_ultima_compra', 'Mediana', 'Media']

resultados = resultados.reindex(columns=nova_ordem)


resultados['Variacao_%'] = ((resultados['Valor_Maximo'] - resultados['Valor_Minimo']) / resultados['Valor_Minimo']) * 100


formato_data_hora = "%Y%m%d_%H%M%S"
data_hora_atual = datetime.datetime.now().strftime(formato_data_hora)

nome_arquivo = f"Variação_Custo_2024_pelo_valor_pago{data_hora_atual}.xlsx"

caminho_saida = "Downloads/"

resultados.to_excel(caminho_saida + nome_arquivo, index=False)

print(df.head(10))
