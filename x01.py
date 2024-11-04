import pandas as pd
import pyautogui
import time
import webbrowser
import math

url = "https://admin.bluesundobrasil.com.br/pedidos"  # URL deve ser uma string
quantidade = 1
loadscreen = 4

def imprimir(id):
    # Acessa a URL
    pyautogui.moveTo(x=440, y=64, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')
    pyautogui.typewrite(url)  # Escreve a URL
    pyautogui.hotkey('enter')

    # Escreve o ID
    time.sleep(loadscreen) #aguarda a pagina
    pyautogui.moveTo(x=368, y=447, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')
    pyautogui.typewrite(str(id))  # Converte id para string
    pyautogui.hotkey('enter')
    time.sleep(5)

    # Clica no botao para entrar no pedido
    pyautogui.moveTo(x=246, y=702, duration=0.5)
    pyautogui.click()

    # Clica para acessar o check list
    time.sleep(loadscreen) #aguarda a pagina
    pyautogui.moveTo(x=440, y= 474, duration=0.5)
    pyautogui.click()

    # Realiza a impressão
    time.sleep(loadscreen) #aguarda a pagina
    pyautogui.moveTo(x=1145, y=283, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'p')
    time.sleep(0.75)
    pyautogui.moveTo(x=1145, y=283, duration=0.5)
    pyautogui.click()
    time.sleep(0.25)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('backspace')
    pyautogui.typewrite(str(quantidade))
    time.sleep(0.75)
    pyautogui.hotkey('enter')
    time.sleep(0.75)

def main():
    excel_file = 'excel.xlsx'
    df = pd.read_excel(excel_file, usecols=[0])  # Le apenas a primeira coluna

    if df.shape[0] < 1:  # Verifica se há pelo menos uma linha
        print("O arquivo deve conter pelo menos uma linha.")
        return
    
    num_linhas = len(df)  # Pega a quantidade de linhas

    for index in range(num_linhas):
        id = df.iloc[index, 0]  # Pega o valor da primeira coluna
        imprimir(id)

if __name__ == "__main__":
    main()
