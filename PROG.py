import pyautogui
import pandas as pd
import time
import tkinter as tk
from tkinter import messagebox

# Carregar a planilha Excel
df = pd.read_excel(r'C:\Users\Kaique\Documents\ATUALIZAR PROG.xlsx')

# Coordenadas dos campos
coordenadas_codigo = (171, 183)  # Coordenadas do campo de código
coordenadas_gestor = (199, 597)  # Coordenadas do campo de gestor
coordenadas_motorista_fixo = (149, 403)  # Coordenadas da caixa "motorista fixo"

# Cor do pixel quando o checkbox está marcado
cor_checkbox_marcado = (0, 0, 0)  # Cor preta #000000

# Função para verificar a cor do checkbox
def verificar_checkbox_marco():
    cor_pixel = pyautogui.pixel(coordenadas_motorista_fixo[0], coordenadas_motorista_fixo[1])
    if cor_pixel == cor_checkbox_marcado:
        print("Checkbox está marcado.")
        return True
    else:
        print("Checkbox está desmarcado.")
        return False

# Função para aguardar e localizar um elemento na tela
def aguardar_localizar_imagem(imagem, timeout=30):
    start_time = time.time()
    pos = None
    while pos is None:
        try:
            pos = pyautogui.locateOnScreen(imagem, confidence=0.7, grayscale=True)
        except Exception as e:
            print(f"Erro ao localizar a imagem: {e}")
        if time.time() - start_time > timeout:
            raise Exception(f"Imagem {imagem} não encontrada dentro do tempo limite.")
    return pos

def exibir_janela_concluido():
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal do Tkinter
    messagebox.showinfo("Processo Concluído", "O processo foi concluído com sucesso!")
    root.destroy()

# Caminho das imagens
caminho_imagens = r'C:\Users\Kaique\Documents\codigos\automacao_datapar'

# Localizar a imagem do campo "Código" apenas uma vez antes do loop
try:
    aguardar_localizar_imagem(f'{caminho_imagens}\\campo_codigo.png')
    print("Campo 'Código' encontrado. Iniciando o processo.")
except Exception as e:
    print(e)
    print("Campo 'Código' não encontrado. O processo será interrompido.")
    exit()  # Sair se o campo não for encontrado

# Loop pelas linhas da planilha
for index, row in df.iterrows():
    # Converter para string e garantir que não sejam nulos
    codvei = str(row['CODVEI']).strip() if pd.notnull(row['CODVEI']) else ''
    prog = str(row['PROG']).strip() if pd.notnull(row['PROG']) else ''

    # Verificar se codvei e prog não estão vazios
    if codvei and prog:
        # Clicar no campo "Código" usando coordenadas
        pyautogui.click(coordenadas_codigo)

        # Esperar um momento para garantir que o campo esteja ativo
        time.sleep(0.10)

        # Simular dois cliques no campo "Código"
        pyautogui.doubleClick()

        # Esperar um momento antes de deletar o conteúdo
        time.sleep(0.10)

        # Pressionar 'delete' para limpar o campo
        pyautogui.press('delete')

        # Pressionar 'backspace' várias vezes para garantir que o campo esteja vazio
        pyautogui.press('backspace', presses=10)

        # Digitar o novo código
        pyautogui.typewrite(codvei)
        pyautogui.press('enter')

        # Aguardar 5 segundos antes de preencher o nome do gestor
        time.sleep(5)

        # Clicar no campo "Gestor" usando coordenadas
        pyautogui.click(coordenadas_gestor)

        # Digitar o nome do gestor
        pyautogui.typewrite(prog)

        # Aguardar 10 segundos antes de verificar a caixa "motorista fixo"
        time.sleep(10)

        # Verificar o estado da caixa "motorista fixo"
        if verificar_checkbox_marco():
            # Se a caixa estiver marcada, clicar para desmarcar
            pyautogui.click(coordenadas_motorista_fixo)
            time.sleep(1)  # Esperar um momento para garantir que a ação foi concluída

        # Localizar e clicar no botão "Salvar"
        try:
            botao_salvar = aguardar_localizar_imagem(f'{caminho_imagens}\\botao_salvar.png')
            pyautogui.click(botao_salvar)
        except Exception as e:
            print(e)
            print("Erro ao localizar o botão 'Salvar'.")

        # Aguardar 5 segundos para garantir que o sistema processe o salvamento
        time.sleep(5)

        # Pressionar 'Enter' após salvar
        pyautogui.press('enter')

        # Aguardar mais 5 segundos
        time.sleep(5)

        # Pressionar 'Enter' novamente
        pyautogui.press('enter')

        # Voltar a clicar no campo "Código" para iniciar o processo novamente
        pyautogui.click(coordenadas_codigo)

# Exibir a janela de conclusão
exibir_janela_concluido()

print("Processo concluído.")
