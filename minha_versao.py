import openpyxl
import pyperclip # uma forma de copiar as informações com acentos, bóne
import pyautogui #movimentar o mouse
from time import sleep

workbook = openpyxl.load_workbook('Y:\Scripts\python\Projetos\\automacao_800reais\\automacao_excel\produtos_ficticios.xlsx')
# produtos é o nome da planilha, se a planilha fosse, folha 1,2,3,4, teria que colocar o nome da sheet da folha
sheet_produtos = workbook['Produtos']


# min_row, é comece a pegar os dados da linha 2
for linha in sheet_produtos.iter_rows(min_row=2):
    # por que 0? na linha o valor 0, corresponde ao nome do produto, e o valor 1, corresponde a descrição do item.
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(44,174,duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(41,250,duration=1)
    pyautogui.hotkey('ctrl','v')


    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(44,395,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(92,481,duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(75,569,duration=1)
    pyautogui.hotkey('ctrl','v')

    
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(77,651,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(89,706,duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)    
    pyautogui.click(44,196,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(45,273,duration=1)
    pyautogui.hotkey('ctrl','v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(46,359,duration=1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(47,441,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(51,534,duration=1)
    tamanho = linha[10].value
    if tamanho == 'Pequeno':
        pyautogui.click(69,577,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(106,599,duration=1)
    else:
        pyautogui.click(91,620,duration=1)

    
    #se for pequeno clicar em uma posição
    #se for médio clicar em uma posição
    #se for grande clicar em uma posição
    # pyperclip.copy(tamanho)
    # pyautogui.click(44,196,duration=1)
    # pyautogui.hotkey('ctrl','v')

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(123,627,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(92,677) #Botão proximo
    
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(74,221,duration=1)
    pyautogui.hotkey('ctrl','v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(84,302,duration=1)
    pyautogui.hotkey('ctrl','v')

    obsevacoes = linha[14].value
    pyperclip.copy(obsevacoes)
    pyautogui.click(67,402,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(76,526,duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(64,612,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(90,668,duration=1)#botão concluir
    pyautogui.click(677,168,duration=1)#botão OK
    sleep(3)
    pyautogui.click(510,451,duration=1)#botão novo produto
    sleep(3)
    

    
