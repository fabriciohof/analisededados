import pandas as pd 
import openpyxl
import smtplib
import email.message
import datetime
from datetime import date

#importando a tabela com excel
tabela_Vendas = pd.read_excel('vendas.xlsx')

#mostrar todas as colunas da tabela
pd.set_option('display.max_columns',None)
# print(tabela_Vendas)

#Mostrar apenas as colunas da Loja e seu Valor Final

# tabela_Vendas[["ID Loja","Valor Final"]]

#Agrupamento com as lojas e a soma dos valores por Loja

# tabela_Vendas.groupby("id Loja").sum()

total = tabela_Vendas[["ID Loja","Valor Final"]].groupby("ID Loja").sum()
print(total)

print ("-" * 55)

#Quantidade total de produtos vendidos por Loja
qtd = tabela_Vendas[["ID Loja","Quantidade"]].groupby("ID Loja").sum()
print(qtd)

print ("-" * 55)

#Resultado de quanto a Loja faturou pela quantidade de vendas que ela teve
#Em média quanto que custou um produto que uma Loja vendeu?

media = (total["Valor Final"] / qtd["Quantidade"]).to_frame()
media = media.rename(columns={0: "Média"})
# to_frame é para converter a divisão de colunas para uma tabela
print(media)

print("-" * 45)
data_atual = date.today()
data_atual = data_atual.strftime('%d/%m/%Y')
print("Relatório do dia:")
print(data_atual)

#Enviar um relatório pelo gmail

import smtplib
import email.message

def enviar_email():  
    corpo_email = f"""
    <p> <b>Relatório das vendas </b> </p>
    <p> Visualize abaixo </p>

    <p> <b>Faturamento:</b> </p>
    {total.to_html(formatters={"Valor Final":"R${:,.0f}".format})}
    <p> <b>Quantidade vendida:</b> </p>
    {qtd.to_html()}
    <p> <b>Média dos Produtos da Loja:</b> </p>
    {media.to_html(formatters={"Média":"R${:,.2f}".format})}
   
    <p> Att.  : Formentini</p>
    """

    msg = email.message.Message()
    msg['Subject'] = "Assunto"
    msg['From'] = 'fabricioformentini2@gmail.com'
    msg['To'] = 'formentinipython@gmail.com'
    password = 'rlnfwcmadcdtotwh' 
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    
    
enviar_email()

print('Email enviado')