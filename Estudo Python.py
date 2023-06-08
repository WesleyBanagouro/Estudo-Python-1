import pandas as pd

import win32com.client as win32

# Importar a base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None)


# Faturamento por loja
Faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()


# Quantidade de produtos vendidos por loja
Quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(Quantidade)

# Ticket médio por produto em cada loja
Ticket_médio = (Faturamento['Valor Final'] / Quantidade['Quantidade']).to_frame()
Ticket_médio = Ticket_médio.rename(columns={0: 'Ticket Médio'})
print(Ticket_médio)

# Enviar e-mail com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'wbanagouro@gmail.com'
mail.Subject = 'Relatório de vendas por loja'
mail.HTMLBody = f'''
<p>Bom dia, prezados,</p>

<p>Segue abaixo tabela resumida de faturamento por cada loja, ticket médio por produto e quantidade de produtos vendidos por cada loja.</p>

<p>Faturamento por cada loja:</p>
<p>{Faturamento.to_html(formatters={'Valor Final': 'R$ {:,.2f}'.format})}</p>

<p>Ticket médio por produto:</p>
<p>{Ticket_médio.to_html(formatters={'Ticket Médio': 'R$ {:,.2f}'.format})}</p>

<p>Quantidade de produtos vendidos:</p>
<p>{Quantidade.to_html()}</p>

<p>Qualquer duvida fico à disposição.</p>

<p>att</p>

'''

mail.Send()

print('=' * 50)

print('Email enviado')
