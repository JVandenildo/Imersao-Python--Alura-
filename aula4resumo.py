import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import mplfinance as mpf
import yfinance as yf
import plotly.graph_objects as go
from plotly.subplots import make_subplots

dados = yf.download('PETR4.SA', start='2023-01-01',end='2023-12-31')
# print(dados)

dados.columns = ['Abertura', 'Maximo', 'Minimo', 'Fechamento', 'Fech_Ajust', 'Volume']
dados = dados.rename_axis('Data')
# print(dados)

dados['Fechamento'].plot(figsize=(10,6))
plt.title('Variação do preço por data', fontsize=16)
plt.legend(['Fechamento'])

df = dados.head(60).copy()
# convertendo o índice em uma coluna de data
df['Data'] = df.index
# convertendo as datas para o formato numérico de matplotlib
# isso é importante para o matplotlib conseguir plotar as datas corretamente no gráfico
df['Data'] = df['Data'].apply(mdates.date2num)
# print(df)

fig, ax = plt.subplots(figsize=(15, 8))

width = 0.7

for i in range(len(df)):
    if df['Fechamento'].iloc[i] > df['Abertura'].iloc[i]:
        color = 'green'
    else:
        color = 'red'
    
    # desenhando a linha vertical do candle (mecha)
    # essa linha mostra os preços máximo (topo da linha) e mínimo (base da linha) do dia
    # é usado `ax.plot` para desenhar uma linha vertical
    # [df['Data].iloc[i], df['Data].iloc[i]] define o ponto x da linha (a data), e [df['Minimo].iloc[i], df['Maximo].iloc[i]] define a largura da linha
    ax.plot([df['Data'].iloc[i], df['Data'].iloc[i]],
            [df['Minimo'].iloc[i], df['Maximo'].iloc[i]],
            color=color,
            linewidth=1)
    
    ax.add_patch(plt.Rectangle((df['Data'].iloc[i] - width/2, min(df['Abertura'].iloc[i], df['Fechamento'].iloc[i])),
                 width,
                 abs(df['Fechamento'].iloc[i] - df['Abertura'].iloc[i]),
                 facecolor=color))

df['MA7'] = df['Fechamento'].rolling(window=7).mean()
df['MA14'] = df['Fechamento'].rolling(window=14).mean()

# plotando as médias móveis
ax.plot(df['Data'], df['MA7'], color='orange', label='Média móvel 7 dias')
ax.plot(df['Data'], df['MA14'], color='yellow', label='Média móvel 14 dias')
# adicionando as legendas paras as médias móveis
ax.legend()

# formatando o eixo x para mostrar as datas
# configuramos o formato da data e a rotação para melhor legibilidade
ax.xaxis_date()
ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
plt.xticks(rotation=45)

plt.title("Gráfico de Candlestick – PETR4.SA com matplotlib")
plt.xlabel("Data")
plt.ylabel("Preço")

plt.grid(True)

plt.show()

fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                    vertical_spacing=0.1,
                    subplot_titles=('Candlesticks', 'Volume Transacionado'),
                    row_width=[0.2, 0.7])

fig.add_trace(go.Candlestick(x=df.index,
                             open=df['Abertura'],
                             high=df['Maximo'],
                             low=df['Minimo'],
                             close=df['Fechamento'],
                             name='Candlestick'),
                             row=1, col=1)

fig.add_trace(go.Scatter(x=df.index,
                         y=df['MA7'],
                         mode='lines',
                         name='MA7 – Média Móvel 7 Dias'),
                         row=1, col=1)

fig.add_trace(go.Scatter(x=df.index,
                         y=df['MA14'],
                         mode='lines',
                         name='MA14 – Média Móvel 14 Dias'),
                         row=1, col=1)

fig.add_trace(go.Bar(x=df.index,
                     y=df['Volume'],
                     name='Volume'),
                     row=2, col=1)

fig.update_layout(yaxis_title='Preço',
                  xaxis_rangeslider_visible=False,  # Desativa o range slider
                  width=1100, height=600)

fig.show()

dados = yf.download('PETR4.SA', start='2023-01-01', end='2023-12-31')
mpf.plot(dados.head(30), type='candle', figsize = (16,8), volume=True, mav=(7,14))
