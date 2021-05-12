from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from IPython.display import display
import pandas as pd
import win32com.client as win32
##########mostrar tabela completa no display##############
pd.options.display.max_columns = None
pd.options.display.max_rows = None
############################################CÓDIGO FUNÇÕES ESSENCIAIS
#navegador a ser pego#
navegador = webdriver.chrome()
#navegador.get é o site a ser colocado no link de site (ué)#
navegador.get('https://www.google.com.br
#decide o que o navegador irá pegar ou clickar ou seja lá o que for#
navegador.find_element_by_xpath ('...') <aqui ele vai onde é o xpath// .send_keys('texto')ou send_keys(keys.ENTER) # para dar um enter
#send keys é para digitar dentro do espaço selecionado
#criar uma variável para o que for guardado de informação
ex: cotacao = navegador.find_element_by_xpath('xpath').get_attribute.('x_value')
print (cotacao)

##############parte da tabela no excel##############
tabela = pd.read_excel(r'diretório excel')   << chama o exel
tabela.loc[tabela['titulo do que mudar']=='O que vai ser mudado', "cotação" = (cotacaoeuro)
###########criar arquivo formato excel atualizado
tabela.to_excel('nomeexcel.xlsx', index = False) 
#######testando com email:::::
#criar um email
eemail = email.CreateItem(0)
#configurar as informações do seu email
eemail.To = "para o email:"
eemail.Subject = "Assunto"
eemail.HTMLBody = """
<p>Olá Fernando</p>
<p>Tudo bem contigo?</p>
<p>Que bom, gostoso</p>
"""
########anexo##########
anexo = ('C:\Semanapython\nomeexcel.xlsx')
eemail.Attachments.Add(anexo)
eemail.Send()
