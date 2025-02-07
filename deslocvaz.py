import pyautogui
import pandas as pd
import time
import tkinter as tk
from tkinter import messagebox

# Carregar a planilha Excel
df = pd.read_excel(r'C:\Users\Kaique\Documents\DESLOCVAZ.xlsx')

# Coordenadas dos campos e botões
coordenadas_codigo = (171, 183)  # Coordenadas do campo de código
coordenadas_aba_dados_veiculo = (78, 156)  # Coordenadas da aba Dados do Veículo
coordenadas_aba_parametros_veiculo = (253, 159)  # Coordenadas da aba Parâmetros do Veículo
coordenadas_checkbox = (193, 572)  # Coordenadas do checkbox Não controla deslocamento vazio
coordenadas_motorista_fixo = (149, 403)  # Coordenadas da caixa "motorista fixo"
cor_checkbox_marcado = (0, 0, 0)  # Cor preta #000000

# Função para verificar a cor de um checkbox
def verificar_checkbox(coordenadas):
    cor_pixel = pyautogui.pixel(coordenadas[0], coordenadas[1])
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

# Exibir mensagem de conclusão
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
    codvei = str(row['CODVEI']).strip() if pd.notnull(row['CODVEI']) else ''

    # Verificar se codvei não está vazio
    if codvei:
        # Clicar no campo "Código"
        pyautogui.click(coordenadas_codigo)
        time.sleep(0.1)

        # Limpar o campo e digitar o código
        pyautogui.doubleClick()
        time.sleep(0.1)
        pyautogui.press('delete')
        pyautogui.press('backspace', presses=10)
        pyautogui.typewrite(codvei)
        pyautogui.press('enter')

        # Esperar para garantir que os dados sejam carregados
        time.sleep(5)

        # Verificar o estado do checkbox "motorista fixo"
        if verificar_checkbox(coordenadas_motorista_fixo):
            # Clicar no checkbox para desmarcar
            pyautogui.click(coordenadas_motorista_fixo)
            time.sleep(1)
        # Ir para a aba Parâmetros do Veículo
        pyautogui.click(coordenadas_aba_parametros_veiculo)
        time.sleep(2)

        # Verificar o estado do checkbox "Não controla deslocamento vazio"
        if verificar_checkbox(coordenadas_checkbox):
            # Clicar no checkbox para desmarcar
            pyautogui.click(coordenadas_checkbox)
            time.sleep(1)

            # Localizar e clicar no botão "Salvar"
            try:
                botao_salvar = aguardar_localizar_imagem(f'{caminho_imagens}\\botao_salvar.png')
                pyautogui.click(botao_salvar)
            except Exception as e:
                print(e)
                print("Erro ao localizar o botão 'Salvar'.")

            # Aguardar 5 segundos para garantir que o sistema processe o salvamento
            time.sleep(5)

            # Pressionar Enter após salvar
            pyautogui.press('enter')
            time.sleep(2)

        # Voltar para a aba Dados do Veículo
        pyautogui.click(coordenadas_aba_dados_veiculo)
        time.sleep(2)

# Exibir mensagem de conclusão
exibir_janela_concluido()

print("Processo concluído.")
