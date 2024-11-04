import pandas as pd
import pyautogui
import time
import webbrowser
import math

loadscreen = 15
indice_de_impressao = 1.5

# Função para realizar o processo de impressão
def imprimir(link, quantidade):
    if quantidade < 5:
        tempo_impressao = quantidade * indice_de_impressao
        
        # Busca o link
        time.sleep(0.75)
        pyautogui.moveTo(x=300, y=61, duration=0.5)
        pyautogui.click()
        time.sleep(0.25)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        link = link.replace('.0', '')  # Remove '.0' se for indesejado
        pyautogui.typewrite(link)
        pyautogui.hotkey('enter')
        time.sleep(loadscreen)  # Aguarda o carregamento

        # Função para imprimir um tipo específico
        def imprimir_tipo(qtd):
            pyautogui.moveTo(x=1145, y=283, duration=0.5)
            pyautogui.click()
            pyautogui.hotkey('ctrl', 'p')
            time.sleep(0.75)
            pyautogui.moveTo(x=1145, y=283, duration=0.5)
            pyautogui.click()
            time.sleep(0.25)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.typewrite(str(qtd))
            time.sleep(0.75)
            pyautogui.hotkey('enter')
            time.sleep(tempo_impressao)

        # Imprimir tipos diferentes
        imprimir_tipo(1)  # Cabo
        imprimir_tipo(1)  # String
        imprimir_tipo(3)  # Fotofix
        imprimir_tipo(quantidade)  # Painel

        # Literalmente nada
        #time.sleep(0.50)
        #pyautogui.moveTo(x=1180, y=759, duration=0.5)
        #pyautogui.click()

# Função principal
def main():
    # Lê o arquivo Excel
    excel_file = 'etiquetas.xlsx'
    df = pd.read_excel(excel_file)

    # Verifica se as colunas estão corretas
    if df.shape[1] < 2:
        print("O arquivo deve conter pelo menos duas colunas.")
        return

    # Loop através de todas as linhas
    num_linhas = len(df)
    url_base = "https://admin.bluesundobrasil.com.br/checklist/etiqueta_producao/"

    for index in range(num_linhas):
        id_pedido = df.iloc[index, 0]
        quantidade_total = df.iloc[index, 1]

        quantidade_painel = 4 * (math.ceil(quantidade_total / 36))

        # Gera o link
        link = f"{url_base}{id_pedido}"
        print(f"Acessando o link: {link}")

        # Realiza as impressões
        imprimir(link, quantidade_painel)

    # Salva as alterações de volta no Excel
    df.to_excel(excel_file, index=False)

if __name__ == "__main__":
    main()
