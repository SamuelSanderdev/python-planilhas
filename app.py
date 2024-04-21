# C:\Users\Samuel>python
# Python 3.12.0 (tags/v3.12.0:0fb18b0, Oct  2 2023, 13:03:39) [MSC v.1935 64 bit (AMD64)] on win32
# Type "help", "copyright", "credits" or "license" for more information.
# >>> from mouseinfo import mouseInfo
# >>> mouseInfo()

import openpyxl
import openpyxl.workbook
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
# Copiar a informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1103,350,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1099,436,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1096,567,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1096,657,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1094,739,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    dimencoes = linha[5].value
    pyperclip.copy(dimencoes)
    pyautogui.click(1094,828,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # botão próximo
    pyautogui.click(1109,899,duration=1)
    sleep(3)
    
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1093,393,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    quantidade_estoque = linha[7].value
    pyperclip.copy(quantidade_estoque)
    pyautogui.click(1095,476,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.click(1100,567,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1099,652,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    tamanho = linha[10].value
    pyautogui.click(1110,733,duration=1)
    
    if tamanho == 'Pequeno':
        pyautogui.click(1120,766,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click (1141,784,duration=1)
    else:
       pyautogui.click(1116,809,duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1104,821,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Botão próximo
    pyautogui.click(1129,881,duration=1)
    sleep(2)
    
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1099,406,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1104,501,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1102,583,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    codigo_barra = linha[15].value
    pyperclip.copy(codigo_barra)
    pyautogui.click(1109,717,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    localizacao_armazem = linha[15].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1108,802,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    pyautogui.click(1132,865,duration=1)
    sleep(2)
    pyautogui.click(1601,192,duration=1)
    sleep(2)
    pyautogui.click(1428,625,duration=1)
    
# Repetir esses passos para outros campos até preencher campos daquela página
# Clicar em próxima
# Repetir os mesmos passos e ir para a próxima página (página 2)
# Repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir
# Clicar em ok, para finalizar o processo
# Clicar em ok mais uma vez na mensagem de confirmação de salvamento no banco de dados
# Clicar em "adicionar mais um repetir o processo até finalizar a planilha"

# PyautoGUI(automação de clicks e teclado)
# Openpyxl (Leitura e automação de planilhas)