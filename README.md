<h1 align="center">Bot de e-mail</h1>

>Status: Completo ✔️ 

>Linguagem utilizada: Python 🐍

*Um projeto básico capaz de enviar e-mails via outlook de forma automática*

## Como o código funciona
* O código simula o envio de um e-mail contendo uma tabela com informações variadas de determinados produtos. Para testar ou avaliar o código, você pode adicionar seu e-mail pessoal na variável "email". Todo o processo de envio não deve levar mais do que 30 segundos (lembre de verificar a caixa de spam caso não receba nada neste tempo).

```python
import pandas as pd
import win32com.client as win32

tabela = pd.read_excel('Vendas.xlsx')
email = 'guilhermeimprovisado394@gmail.com'

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
mail.To = email
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
```
## Resultado
* Após rodar o programa, o e-mail que será enviado para você terá essas características:
![image](https://user-images.githubusercontent.com/85650237/138314501-08d3e472-e118-4790-8f8c-e826fe9e7247.png)

