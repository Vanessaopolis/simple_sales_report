import pandas as pd
from babel.numbers import format_currency, format_decimal
import win32com.client as win32

# imports the database
table_sales_data = pd.read_excel('sales_data.xlsx')
table_sales_data = table_sales_data.rename(columns={'ID Loja': 'Loja'})

# shows the database to work with
pd.set_option('display.max_columns', None)
print(table_sales_data)

# gets revenue per store
revenue_per_store = table_sales_data[['Loja', 'Valor Final']].groupby('Loja').sum()

# gets the amount of sales per store
sales_per_store = table_sales_data[['Loja', 'Quantidade']].groupby('Loja').sum()

# gets average ticket per product in each store
average_ticket_per_product = (revenue_per_store['Valor Final'] / sales_per_store['Quantidade']).to_frame('Ticket '
                                                                                                         'Médio')

# shows all the data
print(revenue_per_store)
print(sales_per_store)
print(average_ticket_per_product)

# connects Outlook app with python
outlook = win32.Dispatch('outlook.application')

# creates the email object
mail = outlook.CreateItem(0)

# defines the email that will receive the report
mail.To = 'alanandrade.vanessa@gmail.com;matheushenriiqu3@gmail.com'

# defines the email subject
mail.Subject = 'Relatório de Vendas'

# formats the email body
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Como solicitado, segue os relatórios de vendas, por cada loja.<br>
Foram analisados o faturamento total, o volume de venda e o ticket médio, para cada uma das lojas.<br>
Se surgir qualquer dúvida, estou à disposição.<br>
As tabelas seguem abaixo:</p>

<p>Faturamento Total de cada Loja</p>
{revenue_per_store.style.format(lambda v: format_currency(v, 'BRL', locale='pt_BR'), 
                                subset=['Valor Final']).to_html()}

<p>Volume de Venda de cada Loja</p>
{sales_per_store.style.format(lambda v: format_decimal(v, locale='pt_BR'), subset=['Quantidade']).to_html()}

<p>Ticket Médio de cada Loja</p>
{average_ticket_per_product.style.format(lambda v: format_currency(v, 'BRL', locale='pt_BR'), 
                                         subset=['Ticket Médio']).to_html()}

<p>Atenciosamente,<br>
Alana Vanessa Andrade</p>
'''

# sends the email
mail.Send()
