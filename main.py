import pandas as pd
import openpyxl


table_sales = pd.read_excel('base-de-dados.xlsx')
pd.set_option('display.max_columns', None)
print(table_sales)


billing = table_sales[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(billing)

quantity = table_sales[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantity)

ticket_medio = (billing['Valor Final'] / quantity['Quantidade']).to_frame()
print(ticket_medio)
