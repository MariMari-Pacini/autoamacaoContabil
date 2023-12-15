import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    # Nome do Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(875,228,duration=1)
    pyautogui.hotkey('ctrl','v')

    # Descrição
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(840,313,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(824,428,duration=1)
    pyautogui.hotkey('ctrl','v')

    # Codigo Produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(818,505,duration=1)
    pyautogui.hotkey('ctrl','v')

    # Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(800,587,duration=1)
    pyautogui.hotkey('ctrl','v')

    # Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(802,660,duration=1)
    pyautogui.hotkey('ctrl','v')

    # Botão próximo
    pyautogui.click(814,705,duration=1)
    sleep(3)

    # Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(816,250,duration=1)
    pyautogui.hotkey('ctrl','v')

    # Quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(846,323,duration=1)
    pyautogui.hotkey('ctrl','v')

    # Data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(844,411,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(817,488,duration=1)
    pyautogui.hotkey('ctrl','v')

    # Tamanho
    tamanho = linha[10].value
    pyautogui.click(858,557,duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(856,590,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(848,610,duration=1)
    else:
        pyautogui.click(839,626,duration=1)

    # material    
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(832,641,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    # Botão próximo
    pyautogui.click(808,695,duration=1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(885,264,duration=1)
    pyautogui.hotkey('ctrl','v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(866,348,duration=1)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(861,429,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(852,544,duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(835,615,duration=1)
    pyautogui.hotkey('ctrl','v')

    # Botão concluir
    pyautogui.click(813,672,duration=1)
    
    # Botão confirmar
    pyautogui.click(1189,191,duration=1)
    
    # Botão segunda confirmação
    pyautogui.click(1193,189,duration=1)
    
    # iniciar cadastro novamente
    pyautogui.click(1030,482,duration=1)
