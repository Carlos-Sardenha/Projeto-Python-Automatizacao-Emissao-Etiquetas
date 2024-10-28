import pandas as pd
import pyautogui
import time
import webbrowser
import math
import tkinter as tk

tempo_espera = 6

# Função para realizar o processo de impressão
def imprimir(link, quantidade):
    if quantidade < 5:
        indice_de_tempo = quantidade * 0.5

        time.sleep(0.75)
        pyautogui.moveTo(x=300, y=61, duration=0.5)
        pyautogui.click()
        time.sleep(0.25)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        link = link.replace('.0', '')  # Remove '.0' se for indesejado
        pyautogui.typewrite(link)
        pyautogui.hotkey('enter')
        time.sleep(tempo_espera) # aguarda o carregamento ( min 2.75 )
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
        pyautogui.moveTo(x=1180, y=759, duration=0.5)
        pyautogui.click()
        time.sleep(indice_de_tempo)
    else:
        indice_de_tempo = quantidade * 1.5

        webbrowser.open(link, new=0)  # Abre o link no navegador
        time.sleep(tempo_espera)  # Aguarda o carregamento da página
        pyautogui.click(x=1145, y=283)  # Seleciona a tela
        pyautogui.hotkey('ctrl', 'p')  # Abre o menu de impressão
        time.sleep(1.25)
        pyautogui.click(x=1145, y=283)  # Clica no campo de quantidade de páginas
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'a')  # Seleciona todo o texto
        pyautogui.press('backspace')  # Apaga o texto atual
        pyautogui.typewrite(str(quantidade))  # Insere a nova quantidade
        time.sleep(1.25)
        pyautogui.hotkey('enter')
        time.sleep(1.0)
        pyautogui.click(x=1180, y=759)  # Clica no botão de imprimir
        time.sleep(indice_de_tempo)  # Aguarda a impressão completar

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Botão de Pause")
        self.root.geometry("300x100+0+0")  # Tamanho e posição

        # Adicionando um ícone (substitua 'icone.ico' pelo caminho do seu ícone)
        try:
            self.root.iconbitmap('icone.ico')  # Coloque o caminho correto do ícone
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")

        self.is_paused = False

        self.pause_button = tk.Button(root, text="Pause", command=self.toggle_pause)
        self.pause_button.pack(pady=20)

        self.label = tk.Label(root, text="Programa em execução...")
        self.label.pack(pady=20)

    def toggle_pause(self):
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_button.config(text="Despausar")
            self.label.config(text="Programa pausado.")
        else:
            self.pause_button.config(text="Pause")
            self.label.config(text="Programa em execução...")

# Função principal
def main():
    root = tk.Tk()
    app = App(root)

    # Lê o arquivo Excel
    excel_file = 'etiquetas.xlsx'
    df = pd.read_excel(excel_file)

    #scanf("digite o tempo de espera", tempo_espera)

    # Verifica se as colunas estão corretas
    if df.shape[1] < 2:
        print("O arquivo deve conter pelo menos duas colunas.")
        root.destroy()  # Fecha a janela se o arquivo estiver incorreto
        return

    # Calcula o número de linhas do Excel
    num_linhas = len(df)

    url_base = "https://admin.bluesundobrasil.com.br/checklist/etiqueta_producao/"

    # Loop através de todas as linhas
    for index in range(num_linhas):
        id_pedido = df.iloc[index, 0]
        quantidade_total = df.iloc[index, 1]

        quantidade_painel = 4 * (math.ceil(quantidade_total / 36))

        # Define as quantidades para cada tipo de impressão
        quantidade_fotofix = 3
        quantidade_string = 1
        quantidade_cabo = 1

        # Gera o link
        link = f"{url_base}{id_pedido}"
        print(f"Acessando o link: {link}")

        # Realiza as impressões
        for quantidade, tipo in [(quantidade_cabo, "Cabo"), (quantidade_string, "String"), (quantidade_fotofix, "FotoFix"), (quantidade_painel, "Painel")]:
            while app.is_paused:
                root.update()  # Permite que a interface continue respondendo enquanto pausada
            imprimir(link, quantidade)

        # Adiciona comentário na terceira coluna
        comentario = f"id_pedido = {id_pedido} | quantidade {quantidade_cabo + quantidade_string + quantidade_fotofix + quantidade_painel}."
        df.loc[index, 2] = comentario  # Armazena o comentário na terceira coluna

    # Salva as alterações de volta no Excel
    df.to_excel(excel_file, index=False)

    root.mainloop()  # Mover aqui para garantir que a janela permaneça aberta

if __name__ == "__main__":
    main()
