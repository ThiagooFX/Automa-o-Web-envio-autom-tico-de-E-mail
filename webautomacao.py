#bibliotecas necessárias para esse projeto
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By


#criar um navegador 

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

#pesquisar um link
navegador.get("https://www.google.com/")

#localizar elemento no site
#.send_keys("...") é para escrever algum caractere no elemento que foi selecionado
navegador.find_element(By.XPATH,
    "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input").send_keys("cotação do Dólar")

#.send_keys(Keys.NOME DA TECLA EM MAIÚSCULO) é para pressionar uma tecla no elemento selecionado
navegador.find_element(By.XPATH,
    "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input").send_keys(Keys.ENTER)
#.get_attribute("") é para pegar algum valor específico no elemento
cotacao_dolar = navegador.find_element(By.XPATH,
    '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
print(cotacao_dolar)

navegador.get("https://www.google.com/")
navegador.find_element(By.XPATH,
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação euro")
navegador.find_element(By.XPATH,
    '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

cotacao_euro = navegador.find_element(By.XPATH,
    '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
print(cotacao_euro)


navegador.get("https://www.melhorcambio.com/ouro-hoje")

cotacao_ouro = navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute("value")
cotacao_ouro = cotacao_ouro.replace(",", ".")
print(cotacao_ouro)

navegador.quit()



import pandas as pd

tabela = pd.read_excel("Produtos.xlsx")



# atualizar a cotação
# nas linhas onde na coluna "Moeda" = Dólar
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacao_dolar)
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacao_euro)
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacao_ouro)

# atualizar o preço base reais (preço base original * cotação)
tabela["Preço de Compra"] = tabela["Preço Original"] * tabela["Cotação"]

# atualizar o preço final (preço base reais * Margem)
tabela["Preço de Venda"] = tabela["Preço de Compra"] * tabela["Margem"]

# tabela["Preço de Venda"] = tabela["Preço de Venda"].map("R${:.2f}".format)

print(tabela)


tabela.to_excel("Produtosatualizados.xlsx", index=False)


#Enivar o arquivo por E-mail

import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

try:
    fromaddr = "..."                               #Email remetente
    toaddr = '...'                                 #Email destinatário      
    msg = MIMEMultipart()

    msg['From'] = fromaddr 
    msg['To'] = toaddr
    msg['Subject'] = "..."                              #Assunto

    body = "\n"                                     #Corpo do email 

    msg.attach(MIMEText(body, 'plain'))

    filename = '...'                               #Nome do arquivo

    attachment = open('...','rb')                  #Coloca o nome do arquivo novamente


    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(part)

    attachment.close()

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(fromaddr, "...")                          #Senha do Email(senha de apps)
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    server.quit()
    print('\nEmail enviado com sucesso!')
except:
    print("\nErro ao enviar email")




