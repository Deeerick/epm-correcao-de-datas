import pandas as pd
from datetime import timedelta

# Leitura do Excel com os dados
df = pd.read_excel('chamados.xlsx')

# Convertendo os valores da coluna data
df['Data de Início'] = pd.to_datetime(df['Data de Início'], format='%d/%m/%Y')
df['Término Previsto'] = pd.to_datetime(df['Término Previsto'], format='%d/%m/%Y')
df['Próxima Atualização'] = pd.to_datetime(df['Próxima Atualização'], format='%d/%m/%Y')


def ajustar_data(data):
    if data.weekday() == 5:
        return data - timedelta(days=1)
    elif data.weekday() == 6:
        return data - timedelta(days=2)
    else:
        return data



df['Data de Início'] = df['Data de Início'].apply(ajustar_data)
df['Data de Início'] = df['Data de Início'].dt.strftime('%d/%m/%Y')

df['Término Previsto'] = df['Término Previsto'].apply(ajustar_data)
df['Término Previsto'] = df['Término Previsto'].dt.strftime('%d/%m/%Y')

df['Próxima Atualização'] = df['Próxima Atualização'].apply(ajustar_data)
df['Próxima Atualização'] = df['Próxima Atualização'].dt.strftime('%d/%m/%Y')


# Salvando os valores convertidos
df.to_excel('Chamados-Corrigido.xlsx', index=False)
