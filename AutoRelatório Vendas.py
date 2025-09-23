#!/usr/bin/env python
# coding: utf-8

# AULA 1 DA JORNADA PYTHON 

# In[5]:


import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 6

#Abrindo nova aba
pyautogui.hotkey('ctrl', 't')

#copiando link
pyperclip.copy('https://drive.google.com/drive/u/2/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
#Acessando link
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

#acessando o Exportar
pyautogui.click(x=454, y=324, clicks = 4)

pyautogui.PAUSE = 3
# Selencionando as vendas
pyautogui.click(x=454, y=324)

# Baixando as vendas
pyautogui.click(x=519, y=213)


# In[6]:


#Calcular os Indicadores que sao FATURAMENTO E QUANTIDADE DE PRODUTOS
import pandas as pd

tabela  = pd.read_excel(r'D:\Pastas do usuário\Downloads\Vendas - dez.xlsx')


quantidade = tabela['Quantidade'].sum()
faturamento = tabela['Valor Final'].sum()


# In[7]:


pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/2/#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(7)

# Enviar por email o Resultado 
time.sleep(9)
pyautogui.click(x=93, y=168)

pyautogui.write('pythonimpressionador+diretoria@gmail.com')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.write('Relatorio de Vendas')
pyautogui.hotkey('tab')

texto = f''' 

Prezados, Bom Dia!

O Faturamento de ontem foi: R${faturamento:,.2f}
A Quantidade de produtos foi: {quantidade:,}

abs
Pedro Ezequiel
            '''
pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')

pyautogui.press('enter')

