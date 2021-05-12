from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from IPython.display import display
import pandas as pd
import win32com.client as win32
##########mostrar tabela completa no display##############
pd.options.display.max_columns = None
pd.options.display.max_rows = None
##################Dolar
navegador = webdriver.Chrome()
navegador.get('https://www.google.com.br')
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação dolar')
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao = navegador.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacao)
############Euro
siteeuro = ('https://www.google.com/search?q=cota%C3%A7%C3%A3o+euro&rlz=1C1ISCS_pt-PTBR951BR951&oq=cota%C3%A7%C3%A3o+euro&aqs=chrome..69i57.2172j0j4&sourceid=chrome&ie=UTF-8')
navegador.get(siteeuro)
cotacaoeuro = navegador.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotacaoeuro)
######Ouro
siteouro = ('https://www.melhorcambio.com/ouro-hoje#:~:text=O%20valor%20do%20grama%20do,em%20R%24%20304%2C29.')
navegador.get(siteouro)
cotacaoouro = navegador.find_element_by_xpath('//*[@id="comercial"]').get_attribute('value')
cotacaoouro = cotacaoouro.replace(',','.')
print(cotacaoouro)
###########tabela corrigida
tabela = pd.read_excel(r'C:\Users\Ferna\Desktop\aula4\Produtos.xlsx')
tabela.loc[tabela['Moeda']=='Euro', "Cotação"] = float(cotacaoeuro)
tabela.loc[tabela['Moeda']=='Dólar', "Cotação"] = float(cotacao)
tabela.loc[tabela['Moeda']=='Ouro', "Cotação"] = float(cotacaoouro)
tabela['Preço Base Reais'] = tabela['Preço Base Original'] * tabela['Cotação']
tabela['Preço Final'] = tabela['Preço Base Reais'] * tabela['Ajuste']
tabela['Preço Final'] = tabela['Preço Final'].map('{:.2f}'.format)
display(tabela)
###########criar arquivo formato excel atualizado
tabela.to_excel('Produtoss.xlsx', index = False)
#######testando com email:::::

#cria integração com outlook
email = win32.Dispatch('outlook.application')
#criar um email
eemail = email.CreateItem(0)
#configurar as informações do seu email
eemail.To = "annadoliveira09@gmail.com"
eemail.Subject = "É isso po"
eemail.HTMLBody = """
<p>Olá Fernando</p>
<p>Tudo bem contigo?</p>
<p>Que bom, gostoso</p>
"""
anexo = ('C:\Semanapython\Produtoss.xlsx')
eemail.Attachments.Add(anexo)
eemail.Send()
print('enviado')