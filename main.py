import pandas as pd
import openpyxl
import smtplib
import email

table_sales = pd.read_excel('base-de-dados.xlsx')
pd.set_option('display.max_columns', None)
print(table_sales)


billing = table_sales[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(billing)

quantity = table_sales[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantity)

ticket_medio = (billing['Valor Final'] / quantity['Quantidade']).to_frame()
print(ticket_medio)


#Se estiver usando o gmail deve primeiro habilitar o acesso de apps emnos seguros

def send_email():
    email_body = f"""
    <p>Email automatico para seu chefe com o relatorio das vendas.</p>
    
    <p>Faturamento:</p>
    {billing.to_html()}

    <p>Quantidade Vendida:</p>
    {quantity.to_html()}

    <p>Ticket Médio dos Produtos:</p>
    {ticket_medio.to_html()}

    <p>Atenciosamente seu bot favorito</p>
    """

    msg = email.message.Message()
    msg['Subject'] = 'Email automatico'
    msg['From'] = 'Seu email aqui'
    msg['to'] = 'Email de destino aqui.'
    password = 'Sua senha aqui'
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(email_body)

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')
send_email()