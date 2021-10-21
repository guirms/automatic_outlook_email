import pandas as pd
import win32com.client as win32

tabela = pd.read_excel('Vendas.xlsx')


faturamento = tabela[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print('\n', ' ' * 12, 'FATURAMENTO')
print(faturamento)
print('\n\n')

print(' ' * 12,'QTD DE PRODUTOS')
produtostot = tabela[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(produtostot)
print('\n\n')

print(' ' * 13, 'TICKET MÉDIO')
ticket = (faturamento['Valor Final'] / produtostot['Quantidade']).to_frame()
ticket = ticket.rename(columns={0: 'Ticket Médio'})
print(ticket)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'guilhermeimprovisado394@gmail.com'
mail.Subject = 'Message subject'
mail.Body = 'Message body'
mail.HTMLBody = (f''' 
<p>Segue e-mail com as informações solicitadas:</p>

<p>Tabela faturamento:</p>
{faturamento.to_html(formatters={'Valor Final' : 'R${:,.2f}'.format})}


<p>Tabela qtd vendida:</p>
{produtostot.to_html(formatters={'produtostot' : '{:,3f}'.format})}

<p>Tabela ticket médio dos produtos:</p>
{ticket.to_html(formatters={'Ticket' : 'R${:,.2f}'.format})}

''')

mail.Send()

print('\n\nE-mail enviado!')

