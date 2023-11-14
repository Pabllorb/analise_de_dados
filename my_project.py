import pandas as pd
import win32com.client as win32

#imprtar a base de dados
sales_table = pd.read_excel('Vendas.xlsx')


#visualizar a base de dados
#visualizar todas as colunas da base de dados com (pd.set_option(opção, valor))

pd.set_option('display.max_columns', None)
print(sales_table)
print('-' * 50)

#faturamento por loja
#filtro de colunas ID Loja e valor final, agrupando todas as lojas e somando(.sum) o valor final por loja
invoicing = sales_table[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

print(invoicing)
print('-' * 50)
#quantidade de produtos vendidos por loja
#filtra as colunas ID Loja e Quantidade somando(.sum) a quantidade
amount = sales_table[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(amount)

print('-' * 50)
#ticket médio por loja = faturamento / quantidade de produtos vendidos

average_ticket = (invoicing['Valor Final'] / amount['Quantidade']).to_frame()
print(average_ticket)

#Enviar um email com o relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'email@email.com'
mail.Subject = 'Relatório de Vendas'
mail.HTMLBody = '''
Prezados,

Segue o relatório de vendas por loja:

Faturamento:
{}

Quantidade de vendas:
{}

Ticket médio por produto:
{}

Att., 
Pablo
'''
mail.Send()


