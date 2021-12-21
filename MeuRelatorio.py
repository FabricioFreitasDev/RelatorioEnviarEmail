import pandas as pd
import win32com.client as win32

# importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')


print('-' *50)
# vizualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

print('-' *50)
# faturamento por loja
faturamento = tabela_vendas [['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print('-' *50)
# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' *50)
# ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar um email com o relátorio
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'Digite seu email'
mail.Subject = 'Relatório do mês Dezembro - Por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório referente ao mês de Dezembro.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R$:{:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R$:{:,.2f}'.format})}

<p>Qualquer Dúvida Estou à Disposição.</p>

<p>Att.,</p>
<p>Fabricio Freitas</p>
'''

mail.Send()

print('E-Mail Enviado com Sucesso!!!')