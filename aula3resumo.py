import pandas as pd
import plotly.express as px

df_principal = pd.read_excel("acoes_pura.xlsx", sheet_name="Principal")
# print(df_principal)

df_acoesTotal = pd.read_excel("acoes_pura.xlsx", sheet_name="Total_de_acoes")
# print(df_acoesTotal)

df_ticker = pd.read_excel("acoes_pura.xlsx", sheet_name="Ticker")
# print(df_ticker)

df_abaGPT = pd.read_excel("acoes_pura.xlsx", sheet_name="ChatGPT")
# print(df_abaGPT)

df_principal = df_principal[['Ativo', 'Data', 'Último (R$)', 'Var. Dia (%)']].copy()
# print(df_principal)

df_principal = df_principal.rename(columns={'Último (R$)':'Valor_Final','Var. Dia (%)':'Var_Dia_pct'}).copy() # renomear as colunas para retirar os caracteres especiais
# print(df_principal)

df_principal['Var_pct'] = df_principal['Var_Dia_pct'] / 100
df_principal['Var_Inicial'] = df_principal['Valor_Final'] / (1 + df_principal['Var_pct'])
# print(df_principal)

df_principal = df_principal.merge(df_acoesTotal, left_on='Ativo', right_on='Código', how='left')
# print(df_principal)

df_principal = df_principal.drop(columns=['Código'])
# print(df_principal)

df_principal['Variacao_RS'] = (df_principal['Valor_Final'] - df_principal['Var_Inicial'])*df_principal['Qtde. Teórica']
# print(df_principal)

pd.options.display.float_format = '{:.2f}'.format # comando para ajustar como os números são mostrados
df_principal['Qtde. Teórica'] = df_principal['Qtde. Teórica'].astype(int)
# print(df_principal)

df_principal = df_principal.rename(columns={'Qtde. Teórica':'Qtde_Teorica'}).copy() # renomear a coluna para retirar os caracteres especiais
# print(df_principal)

df_principal['Resultado'] = df_principal['Variacao_RS'].apply(lambda x: 'Subiu' if x > 0 else('Desceu' if x < 0 else 'Estável'))
# print(df_principal)

df_principal = df_principal.merge(df_ticker, left_on='Ativo', right_on='Ticker', how='left')
df_principal = df_principal.drop(columns='Ticker')
# print(df_principal)

df_principal = df_principal.merge(df_abaGPT, left_on='Nome', right_on='Empresa', how='left')
df_principal = df_principal.drop(columns='Empresa')
# print(df_principal)

df_principal['Cat_Idade'] = df_principal['Idade (anos)'].apply(lambda x: 'Mais de 100 anos' if x > 100 else('Menos de 50 anos' if x < 50 else 'Entre 50 e 100 anos'))
# print(df_principal)

maior = df_principal['Variacao_RS'].max()
menor = df_principal['Variacao_RS'].min()
media = df_principal['Variacao_RS'].mean()
media_subiu = df_principal[df_principal['Resultado'] == 'Subiu']['Variacao_RS'].mean()
media_desceu = df_principal[df_principal['Resultado'] == 'Desceu']['Variacao_RS'].mean()

print(f'Maior R${maior:,.2f};\nMenor R${menor:,.2f};\nMédia R${media:,.2f};\nMédia subiu R${media_subiu:,.2f};\nMédia desceu R${media_desceu:,.2f}')

df_principal_subiu = df_principal[df_principal['Resultado'] == 'Subiu']
# print(df_principal_subiu)

df_analise_segmento = df_principal_subiu.groupby('Segmento')['Variacao_RS'].sum().reset_index()
# print(df_analise_segmento)

df_analise_saldo = df_principal.groupby('Resultado')['Variacao_RS'].sum().reset_index()
# print(df_analise_saldo)

fig = px.bar(df_analise_saldo, x='Resultado', y='Variacao_RS', text='Variacao_RS', title='Variação R$ por resultado')
fig.show()