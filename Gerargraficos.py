import pandas as pd
import plotly.express as px
import locale

# Configurar a localidade para português do Brasil
locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

df = pd.read_excel('C:\\Users\\patricia.messina\\Desktop\\Dados-de-rescisão.xlsx')

# Exibir as primeiras linhas
print(df.head())

# Verificar se a coluna 'Líquido Rescisão' contém valores numéricos
if pd.api.types.is_numeric_dtype(df['Líquido Rescisão']):
    # Somar os valores da coluna 'Líquido Rescisão'
    soma_liquido_rescisao = df['Líquido Rescisão'].sum()
    print(f'Soma dos valores da coluna "Líquido Rescisão": {soma_liquido_rescisao}')
else:
    print('A coluna "Líquido Rescisão" não contém valores numéricos.')

# Converter a coluna 'Líquido Rescisão' para numérico, se necessário
df['Líquido Rescisão'] = pd.to_numeric(df['Líquido Rescisão'], errors='coerce')

# Agrupar por 'Nome Empresa' e calcular a soma dos valores da coluna 'Líquido Rescisão'
soma_por_nome_empresa = df.groupby('Nome Empresa')['Líquido Rescisão'].sum().reset_index()

# Criar o gráfico de barras usando plotly.express
fig = px.bar(soma_por_nome_empresa, x='Nome Empresa', y='Líquido Rescisão',
             title='Soma do Líquido Rescisão por Nome Empresa',
             labels={'Nome Empresa': 'Nome Empresa'},
             color='Líquido Rescisão',  # Adiciona cores baseadas na soma
             color_continuous_scale='gray')  # Define a paleta de cores

# Exibir o gráfico
fig.show()


df2 = pd.read_excel('C:\\Users\\patricia.messina\\Desktop\\Dados-de-rescisão.xlsx')

agrupar_por_dt_demissao = df2.groupby('Nome Empresa')['Dt Demissão'].count().reset_index()

# Criar o gráfico de barras usando plotly.express
fig2 = px.bar(agrupar_por_dt_demissao, x='Nome Empresa', y='Dt Demissão',
             title='Quantidade de demissões por Nome Empresa',
             labels={'Nome Empresa': 'Nome Empresa', 'Dt Demissão': 'Quantidade de Demissões'},
             color='Dt Demissão',  # Adiciona cores baseadas na quantidade
             color_continuous_scale='magenta')  # Define a paleta de cores

# Exibir o gráfico
fig2.show()


# Carregar o arquivo Excel
df3 = pd.read_excel('C:\\Users\\patricia.messina\\Desktop\\Dados-de-rescisão.xlsx')

# Certificar-se de que a coluna de datas está em formato datetime
df3['Dt Demissão'] = pd.to_datetime(df3['Dt Demissão'], errors='coerce')

# Extrair o mês da data de demissão e adicionar uma coluna com o nome do mês
df3['Mes'] = df3['Dt Demissão'].dt.month
df3['Mes_Descritivo'] = df3['Dt Demissão'].dt.strftime('%B')  # Formato para mês descritivo

# Agrupar por Nome Empresa e Mês Descritivo, e somar o Líquido Rescisão
agrupar_por_mes = df3.groupby(['Nome Empresa', 'Mes_Descritivo'])['Líquido Rescisão'].sum().reset_index()

# Criar o gráfico de barras usando plotly.express
fig3 = px.bar(agrupar_por_mes, x='Nome Empresa', y='Líquido Rescisão',
             title='Líquido Rescisão por Mês e Empresa',
             labels={'Nome Empresa': 'Nome Empresa', 'Líquido Rescisão': 'Soma do Líquido Rescisão', 'Mes_Descritivo': 'Mês'},
             color='Mes_Descritivo',  # Adiciona cores baseadas no mês descritivo
             color_continuous_scale='rainbow')  # Define a paleta de cores

# Exibir o gráfico
fig3.show()