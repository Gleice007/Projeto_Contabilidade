"""

1º Entrar na planilha.

2º Copiar informações de um campo e colar no seu campo correspondente.

3º Repetir o segundo passo para outros campos até preencher os campos 
daquela pagina.

4º Clicar em proximo.

5º Repetir os mesmos passos e ir para a proxima página(página 2).

6º Repetir os mesmos passos e finalizar o cadastro daquele produto e 
clicar em concluir.

7º Clicar em ok, para finalizar o processo.

8º Clicar no ok mais uma vez, na mensagem de confirmação de salvamento
no banco de dados. 

9º Clicar em "adicionar mais um e repetir o processo até finalizar a planilha".
"""

# PyAutoGUI(automação de click e teclado)
# Openpyxl (Leitura e automação de planilhas)


import openpyxl
import pyperclip
import pyautogui
from time import sleep


# 1º Entrar na planilha.
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# 2º Copiar informações de um campo e colar no seu campo correspondente.
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(84,258,duration=1)
    pyautogui.hotkey('ctrl','v')

    descrição = linha[1].value
    pyperclip.copy(descrição)
    pyautogui.click(67,347,duration=1)
    pyautogui.hotkey('ctrl','v')


    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(61,478,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(58,565,duration=1)
    pyautogui.hotkey('ctrl','v')


    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(59,647,duration=1)
    pyautogui.hotkey('ctrl','v')


    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(59,737,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(86,801,duration=1)
    sleep(4)



    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(62,279,duration=1)
    pyautogui.hotkey('ctrl','v')


    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(56,368,duration=1)
    pyautogui.hotkey('ctrl','v')


    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(59,453,duration=1)
    pyautogui.hotkey('ctrl','v')


    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(59,542,duration=1)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyautogui.click(68,626,duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(90,661,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(81,690,duration=1)
    else:
        pyautogui.click(86,724,duration=1)


    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(61,713,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(69,773,duration=1)
    sleep(4)
    
    
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(64,319,duration=1)
    pyautogui.hotkey('ctrl','v')


    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(66,404,duration=1)
    pyautogui.hotkey('ctrl','v')


    observacao = linha[14].value
    pyperclip.copy(observacao)
    pyautogui.click(66,497,duration=1)
    pyautogui.hotkey('ctrl','v')


    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(61,623,duration=1)
    pyautogui.hotkey('ctrl','v')


    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(63,713,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(80,772,duration=1)
    sleep(3)

    pyautogui.click(569,187,duration=1)
    sleep(2)

    pyautogui.click(390,530,duration=1)
    