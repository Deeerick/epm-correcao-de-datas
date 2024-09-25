import pandas as pd
from datetime import timedelta, datetime


# Leitura do Excel com os dados
df = pd.read_excel('Chamados.xlsx')


# Função para ajustar e formatar as datas
def ajustar_e_formatar_data(coluna) -> None:
    df[coluna] = pd.to_datetime(df[coluna], format='%d/%m/%Y')
    df[coluna] = df[coluna].apply(ajustar_data)
    df[coluna] = df[coluna].dt.strftime('%d/%m/%Y')


def ajustar_data(data):
    if data.weekday() == 5:  # Se for sábado
        return data - timedelta(days=1)
    elif data.weekday() == 6:  # Se for domingo
        return data - timedelta(days=2)
    else:
        return data


colunas_para_ajustar = ['Data de Início',
                        'Término Previsto', 'Próxima Atualização']
for coluna in colunas_para_ajustar:
    ajustar_e_formatar_data(coluna)


# Salvando os valores convertidos
data_atual = datetime.now().strftime('%d.%m.%Y')
df.to_excel(f'{data_atual} - Chamados-Corrigido.xlsx', index=False)
print('Processo concluído!')
