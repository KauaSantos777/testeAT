import pyautogui
import time

# Dados para preencher a planilha
produtos = ['Caneta', 'Lapis', 'Caderno', 'Borracha', 'Apontador']
valores = [2, 3, 12.00, 1, 5]
quantidades = [10, 20, 5, 15, 7]
cores =['azul', 'vermelho', 'amarelo', 'verde', 'preto']

#INICIO DO SCRIPT
pyautogui.PAUSE = 0.4

# Passo 1 - Abrir a ferramenta
pyautogui.press('win')
pyautogui.write('Excel')
pyautogui.press('enter')

time.sleep(2)

# Passo 2 - Criar uma nova planilha
pyautogui.click(x=267, y=198)

time.sleep(1)

# Passo 3 - Preencher a planilha com os dados
pyautogui.click(x=112, y=256 ,clicks=2)
pyautogui.write('Produtos')
pyautogui.press('enter')
pyautogui.write('Valores')
pyautogui.press('enter')
pyautogui.write('Quantidades')
pyautogui.press('enter')
pyautogui.write('Cores')


time.sleep(2)

pyautogui.click(x=178, y=257 ,clicks=2)
for produto in produtos:
    pyautogui.write(produto)
    pyautogui.press('tab')
    time.sleep(0.5)

pyautogui.click(x=195, y=277, clicks=2)
for valor in valores:
    pyautogui.write(str(valor))
    pyautogui.press('tab')
    time.sleep(0.5)

pyautogui.click(x=201, y=296, clicks=2)       
for quantidade in quantidades:
    pyautogui.write(str(quantidade))
    pyautogui.press('tab')
    time.sleep(0.5)

pyautogui.click(x=166, y=320, clicks=2)
for cor in cores:
    pyautogui.write(cor)
    pyautogui.press('tab')
    time.sleep(0.5)  

#FIM DA TAREFA