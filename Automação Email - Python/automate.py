#importar as bibliotecas
import pyautogui # faz a automação do mouse e do teclado
import time # controla o tempo do nosso programa
import pyperclip # permite copiar e colar com o python

pyautogui.PAUSE = 1

pyautogui.alert("O programa vai começar, não utilize o computador")
#Passo 1: entrar no sistema(Link do google Drive)
pyautogui.hotkey('win')
pyautogui.write('chrome')
pyautogui.press('enter')
link = 'https://drive.google.com/drive/shared-with-me'
#copia o link
pyperclip.copy(link)
#cola o link
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
# esperar um pouquinho
time.sleep(2)

#Passo 2: entrar na pasta da aula 1
pyautogui.click(x=368, y=332, clicks=2)

#Passo 3: fazer o download da base de vendas
time.sleep(2)
pyautogui.click(x=1124, y=392)
time.sleep(2)
pyautogui.click(x=1712, y=187)
time.sleep(2)
pyautogui.click(x=1597, y=556)
time.sleep(5)

#passo 4: calcular os indicadores (faturamento e quantidade de produtos)
import pandas as pd

tabela = pd.read_excel(r'Vendas - Dez.xlsx')
display(tabela)
faturamento = tabela["Valor Final"].sum()
qtde_produtos = tabela["Quantidade"].sum()

#passo 5: entrar no meu email
#email da diretoria: masj_1+diretoria@outlook.com
pyautogui.hotkey('ctrl','t')
link = 'https://outlook.live.com/mail/0/inbox'
#copia o link
pyperclip.copy(link)
#cola o link
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
time.sleep(5)

#passo 6: criar o email
pyautogui.click(x=195, y=165)
time.sleep(2)
pyautogui.click(x=786, y=229)
time.sleep(1)
pyautogui.write('masj_1+diretoria@outlook.com')
time.sleep(1)
pyautogui.press('tab')

pyautogui.click(x=763, y=287)
assunto = "Relatório de Vendas"
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl','v')


#passo 7: enviar o email
pyautogui.press('tab')
time.sleep(1)
texto_email = f"""
Prezados, bom dia

O faturamento foi: R${faturamento:,.2f}
A quantidade de produtos foi: {qtde_produtos:,}

Abs
Maurício Alves da Silva Júnior
"""
pyperclip.copy(texto_email)
pyautogui.hotkey('ctrl','v')
pyautogui.hotkey('ctrl','enter')

pyautogui.alert("o programa terminou")