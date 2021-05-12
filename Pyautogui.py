import pyperclip
import time
import pyautogui as pag
#import pandas == le excel e os krl a4#
#import os == serve pra quando tem nomes diferentes os arquivos criados#
n = 0
email = input('Digite o email que vai ser enviado: \n')
assunto = input('qual o assunto?: \n')
mensagem = input('qual a mensagem?: \n')
link = 'outlook.com'
pag.press('win')
time.sleep(1)
pag.write('opera')
time.sleep(1)
pag.press('enter')
time.sleep(2)
pyperclip.copy(link)
pag.hotkey('ctrl', 'v')
pag.press('enter')
time.sleep(4)
while n <= 3:
    pag.click(182, 138)
    time.sleep(4)
    pyperclip.copy(email)
    pag.hotkey('ctrl', 'v')
    pag.press('tab')
    pyperclip.copy(assunto)
    pag.hotkey('ctrl', 'v')
    pag.press('tab')
    pyperclip.copy(mensagem)
    pag.hotkey('ctrl', 'v')
    pag.click(354, 650)
    time.sleep(3)
    pag.click(182, 138)
    time.sleep(4)
    pyperclip.copy(email)
    pag.hotkey('ctrl', 'v')
    pag.press('tab')
    pyperclip.copy(assunto)
    pag.hotkey('ctrl', 'v')
    pag.press('tab')
    pyperclip.copy(mensagem)
    pag.hotkey('ctrl', 'v')
    pag.click(354, 650)
    time.sleep(3)
    n = n +1
else:
    print('acabou')

