#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# In[43]:


# Antes de mais nada deve importar as bibliotecas
import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1 # pausa de tempo de um segundo entre cada comando

# passo 1: Entrar no sistema (https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing)
pyautogui.hotkey('ctrl', 't')  # hotkey quando for clicar em mais de uma tecla
link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing" # declaração de uma variavél
pyperclip.copy(link) # para copiar o link com o caracter especial
pyautogui.hotkey('ctrl', 'v')
# pyautogui.write("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing" write para escreve o link ou o que for pra ser digitado porém ele não irá digitar qualquer caracter especial
pyautogui.press("enter") # press quando for somente uma tecla

# passo 2: Navegar no sistema (Entrar na pasta exportar)
time.sleep(10) # para que nesse momento ele espere 10 segundos
pyautogui.click(x=395, y=292, clicks=2) # Comando para dar um clique em um objeto
# para dar um duplo clique é clicks=2, caso queira clicar no botão direito é button=right

# passo 3: Fazer o download do arquivo
time.sleep(5)
pyautogui.click(x=398, y=334)
pyautogui.click(x=1159, y=195)
pyautogui.click(x=1032, y=588)
time.sleep(15)


# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[44]:


# passo 4: Calcular o faturamento e a quantidade de produtos vendido
# caso esteja em outro editor de código lembrar de instalar o pandas e o openpyxl
import pandas as pd

tabela = pd.read_excel(r"C:\Users\breno\Downloads\Vendas-Dez.xlsx") # importar a tabela de dados para o python, lembrar de colocar na frente do caminho do arquivo o r para que ele possa ler o arquivo extamente como ele está na tabela.
faturamento = tabela["Valor Final"].sum() # Forma como é realizado a soma em uma tabela.
quantidade = tabela["Quantidade"].sum()
display(tabela) # exibe os dados da tabela de forma organizada
print(faturamento) #exibe a informação de forma desorganizada
print(quantidade)


# ### Vamos agora enviar um e-mail pelo gmail

# In[45]:


# passo 5: Enviar o e-mail para a diretoria
pyautogui.hotkey('ctrl', 't')
link = "https://mail.google.com/mail/u/0/#inbox"
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press("enter")

# clicar no botão e-mail
time.sleep(20)
pyautogui.click(x=46, y=210)

# escrever para que será o e-mail
time.sleep(10)
pyautogui.write("brenocaval858@gmail.com") # escreve o correio do destinatário
time.sleep(2)
pyautogui.press("tab") # escolhe o e-mail
time.sleep(2)
pyautogui.press("tab") # passar para o campo assunto

# Escrever o assunto do e-mail
time.sleep(2)
Assunto = "Relatório de vendas"
pyperclip.copy(Assunto) # escrever no assunto
pyautogui.hotkey('ctrl', 'v')
pyautogui.press("tab") # para passar para o corpo do e-mail

# escrever o e-mail
texto = f"""
Prezados, bom dia

O faturamento foi de R${faturamento:,.2f}
A quantidade de produtos foi de {quantidade:,}

Abs

Claudiane Gomes
"""
# sempre que quiser colocar o valor de uma variavel dentro do texto
# deve-se colocar o f antes para indicar que o texto terá uma formatação
# para que as variavéis sejam inseridas em um lugar determinado do texto
# ela deve ser colocado entre chaves {}
# para formatar o número deve-se coinsultar a aopostila ou procurar na net
# é muito confuso então não precisa decorar
# para sinalizar a formatação inserir : para colocar o separador de milhar é a ,
# para colocar duas casas decimais para o caso de dinheiro é com .2f

pyperclip.copy(texto) # escrever no assunto
pyautogui.hotkey('ctrl', 'v')

# enviar o e-mail
time.sleep(2)
pyautogui.click(x=840, y=704)


# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# In[46]:


import time
time.sleep(5)
pyautogui.position() # para saber a posição de um objeto na tela com o mouse encima


# In[ ]:




