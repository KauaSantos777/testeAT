import pyautogui
import time

# Dados para preencher a planilha
produtos = ['Caneta', 'Lapis', 'Caderno', 'Borracha', 'Apontador']
valores = [2, 3, 12.00, 1, 5]
quantidades = [10, 20, 5, 15, 7]
cores = ['azul', 'vermelho', 'amarelo', 'verde', 'preto']

#INICIO DO SCRIPT
pyautogui.PAUSE = 0.4

# Passo 1 - Abrir a ferramenta
pyautogui.press('win')
pyautogui.write('Excel')
pyautogui.press('enter')

time.sleep(2)

# Passo 2 - Criar uma nova planilha
# Certifique-se de ter uma imagem chamada 'novo_arquivo.png' representando o botão "Novo arquivo" do Excel
novo_arquivo_btn = pyautogui.locateCenterOnScreen('novo_arquivo.png', confidence=0.8)
if novo_arquivo_btn:
    pyautogui.click(novo_arquivo_btn)
else:
    print("Botão 'Novo arquivo' não encontrado na tela. Verifique a imagem e tente novamente.")
    exit(1)

time.sleep(1)

# Passo 3 - Preencher a planilha com os dados
pyautogui.click(x=112, y=256, clicks=2)
pyautogui.write('Produtos')
pyautogui.press('enter')
pyautogui.write('Valores')
pyautogui.press('enter')
pyautogui.write('Quantidades')
pyautogui.press('enter')
pyautogui.write('Cores')

time.sleep(2)

def preencher_coluna(dados, x, y):
    pyautogui.click(x=x, y=y, clicks=2)
    for item in dados:
        pyautogui.write(str(item))
        pyautogui.press('tab')
        time.sleep(0.5)

preencher_coluna(produtos, 178, 257)
preencher_coluna(valores, 195, 277)
preencher_coluna(quantidades, 201, 296)
preencher_coluna(cores, 166, 320)

time.sleep(0.5)
#FIM DA TAREFA