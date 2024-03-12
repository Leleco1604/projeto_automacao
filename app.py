import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha

workbook = openpyxl.load_workbook("produtos_ficticios.xlsx")
sheet_produtos = workbook["Produtos"]

# copiar infprmações de umm campo e colar em seu campo correspondente
for linha in sheet_produtos.iter_rows(min=2):

    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(834, 206, duration=1)
    pyautogui.hotkey("ctrl", "v")

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(827, 430, duration=1)
    pyautogui.hotkey("ctrl", "v")

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(849, 509, duration=1)
    pyautogui.hotkey("ctrl", "v")

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(855, 600, duration=1)
    pyautogui.hotkey("ctrl", "v")

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(863, 683, duration=1)
    pyautogui.hotkey("ctrl", "v")

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(826, 606, duration=1)
    pyautogui.hotkey("ctrl", "v")

    # Botão proximo
    pyautogui.click(669, 608, duration=1)
    # Segunda Pagina

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(808, 229, duration=1)
    pyautogui.hotkey("ctrl", "v")

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(831, 402, duration=1)
    pyautogui.hotkey("ctrl", "v")

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(823, 483, duration=1)
    pyautogui.hotkey("ctrl", "v")

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(863, 482, duration=1)
    pyautogui.hotkey("ctrl", "v")

    tamanho = linha[10].value
    pyautogui.click(832, 522, duration=1)
    if tamanho == "Pequeno":
        pyautogui.click(813, 603, duration=1)
    elif tamanho == "Médio":
        pyautogui.click(818, 616, duration=1)
    else:
        pyautogui.click(842, 639, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(827, 608, duration=1)
    pyautogui.hotkey("ctrl", "v")

    # botao proximo
    pyautogui.click(813, 671, duration=1)

    fabricante = linha[12].value
    pais_origem = linha[13].value
    observacoes = linha[14].value
    codigo_de_barras = linha[15].value
    localizaao_armazem = linha[16].value


# Repitir esses passos para outros campos até preencher campos daquela página
# Clicar em próxima
# Repetir os mesmos passos e ir para a próxima página(página 2)
# Repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir
# clicar em ok, para finalizar o processo
# Clicar no ok mais um vez na mensagem de confirmação de salvamento no banco de dados
# Clicar em "adicionar mais um e repetir o processo até finalizar a planilha"
