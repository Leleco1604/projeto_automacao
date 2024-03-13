import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha
workbook = openpyxl.load_workbook("produtos_ficticios.xlsx")
sheet_produtos = workbook["Produtos"]
# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    # Nome do Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(780,231, duration=1)
    pyautogui.hotkey("ctrl", "v")

    # Descrição
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(764, 303, duration=1)
    pyautogui.hotkey("ctrl", "v")

    # Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(782,409, duration=1)
    pyautogui.hotkey("ctrl", "v")

    # Codigo Produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(782,478, duration=1)
    pyautogui.hotkey("ctrl", "v")

    # Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(775,550, duration=1)
    pyautogui.hotkey("ctrl", "v")

    # Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(778,616, duration=1)
    pyautogui.hotkey("ctrl", "v")

    # Botão próximo
    pyautogui.click(781,667, duration=1)
    sleep(3)


    # Segunda Pagina
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(765,250, duration=1)
    pyautogui.hotkey("ctrl", "v")

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(769,322, duration=1)
    pyautogui.hotkey("ctrl", "v")

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(773,389, duration=1)
    pyautogui.hotkey("ctrl", "v")

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(771,456, duration=1)
    pyautogui.hotkey("ctrl", "v")

    tamanho = linha[10].value
    pyautogui.click(803,521, duration=1)
    if tamanho == "Pequeno":
        pyautogui.click(778,554, duration=1)
    elif tamanho == "Médio":
        pyautogui.click(779,570, duration=1)
    else:
        pyautogui.click(779,590, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(777,596, duration=1)
    pyautogui.hotkey("ctrl", "v")

    # botao proximo
    pyautogui.click(781,637, duration=1)

    #Pagina 3

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(782,275, duration=1)
    pyautogui.hotkey("ctrl", "v")

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(771,350, duration=1)
    pyautogui.hotkey("ctrl", "v")

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(774,428, duration=1)
    pyautogui.hotkey("ctrl", "v")

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(772,526, duration=1)
    pyautogui.hotkey("ctrl", "v")

    localizaao_armazem = linha[16].value
    pyperclip.copy(localizaao_armazem)
    pyautogui.click(769,598, duration=1)
    pyautogui.hotkey("ctrl", "v")

    # botão proximo
    pyautogui.click(778,638, duration=1)

    # botão de ok, pos proximo
    pyautogui.click(1199,170, duration=1)

    # botão de adicionar mais um
    pyautogui.click(1028,448, duration=1)

# Repitir esses passos para outros campos até preencher campos daquela página
# Clicar em próxima
# Repetir os mesmos passos e ir para a próxima página(página 2)
# Repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir
# clicar em ok, para finalizar o processo
# Clicar no ok mais um vez na mensagem de confirmação de salvamento no banco de dados
# Clicar em "adicionar mais um e repetir o processo até finalizar a planilha"
