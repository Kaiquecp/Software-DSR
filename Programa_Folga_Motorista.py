import os
import sys
import pandas as pd
import pyodbc
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

server = 'server'
database = 'database'
username = 'user'
password = 'password'

conn_str = (
    'DRIVER={ODBC Driver 18 for SQL Server};'
    f'SERVER={server};'
    f'DATABASE={database};'
    f'UID={username};'
    f'PWD={password};'
    'TrustServerCertificate=yes;'
)
conn = pyodbc.connect(conn_str)

query = """
QUERY
"""

df = pd.read_sql(query, conn)

def calcular_folgas(datas):
    folgas = []
    dias_trabalhados = 0
    ultimo_dia = None

    for i in range(len(datas)):
        if ultimo_dia is not None and (datas[i] - ultimo_dia).days > 1:
            while (ultimo_dia + timedelta(days=1)) < datas[i]:
                ultimo_dia += timedelta(days=1)
                folgas.append(ultimo_dia.strftime("%Y-%m-%d"))

            dias_trabalhados = 0

        ultimo_dia = datas[i]
        dias_trabalhados += 1

        if dias_trabalhados == 6:
            folga = ultimo_dia + timedelta(days=1)
            folgas.append(folga.strftime("%Y-%m-%d"))
            dias_trabalhados = 0

    if dias_trabalhados > 0 and dias_trabalhados < 6:
        dias_faltantes = 6 - dias_trabalhados
        folga = ultimo_dia + timedelta(days=dias_faltantes) + timedelta(days=1)
        folgas.append(folga.strftime("%Y-%m-%d"))

    return list(set(folgas))
	
def calcular_proxima_folga(df):
    df['Data Base'] = pd.to_datetime(df['Data Base'])
    df = df.sort_values(['Nome', 'Data Base'])
    folgas = []

    for nome, grupo in df.groupby('Nome'):
        dias_trabalhados = grupo['Data Base'].dt.date.unique()
        dias_trabalhados.sort()

        folgas_motorista = calcular_folgas(dias_trabalhados)

        if folgas_motorista:
            ultima_folga = max(folgas_motorista)
            ultima_folga_formatada = datetime.strptime(ultima_folga, "%Y-%m-%d").strftime("%d/%m/%Y")
            folgas.append({
                'Nome': nome,
                'Próxima Folga': ultima_folga_formatada
            })

    folgas_df = pd.DataFrame(folgas)
    return folgas_df
	
def exportar_excel():
    folgas_df = calcular_proxima_folga(df)
    arquivo_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')])
    if arquivo_path:
        folgas_df.to_excel(arquivo_path, index=False)
        messagebox.showinfo("Exportação", "Arquivo exportado com sucesso!")

def consultar_motoristas():
    motoristas_selecionados = nome_var.get().strip()
    data_inicio = data_inicio_var.get().strip()
    data_fim = data_fim_var.get().strip()

    resultado_textbox.delete('1.0', tk.END)
    total_motoristas = 0

    if motoristas_selecionados:
        motoristas_lista = motoristas_selecionados.split(', ')
        folga_motorista = df_folgas[df_folgas['Nome'].isin(motoristas_lista)]
        total_motoristas = len(folga_motorista)
        if not folga_motorista.empty:
            resultado = ""
            for index, row in enumerate(folga_motorista.iterrows(), 1):
                _, data = row
                resultado += f"{index}. {data['Nome']},\n PRÓXIMA FOLGA: {data['Próxima Folga']}\n\n"
            resultado_textbox.insert(tk.END, resultado)
        else:
            resultado_textbox.insert(tk.END, "Motoristas não encontrados ou sem folgas calculadas\n")
        total_motoristas_label.config(text=f"Total de Motoristas: {total_motoristas}")
        return

    if data_inicio and data_fim:
        try:
            data_inicio_dt = datetime.strptime(data_inicio, "%d/%m/%Y")
            data_fim_dt = datetime.strptime(data_fim, "%d/%m/%Y")
            folgas_periodo = df_folgas[
                (pd.to_datetime(df_folgas['Próxima Folga'], format="%d/%m/%Y") >= data_inicio_dt) &
                (pd.to_datetime(df_folgas['Próxima Folga'], format="%d/%m/%Y") <= data_fim_dt)
            ]
            total_motoristas = len(folgas_periodo)
            if not folgas_periodo.empty:
                resultado = ""
                for index, row in enumerate(folgas_periodo.iterrows(), 1):
                    _, data = row
                    resultado += f"{index}. {data['Nome']},\n PRÓXIMA FOLGA: {data['Próxima Folga']}\n\n"
                resultado_textbox.insert(tk.END, resultado)
            else:
                resultado_textbox.insert(tk.END, "Nenhuma folga encontrada no período especificado\n")
        except ValueError:
            messagebox.showerror("Erro", "Formato de data inválido. Por favor, use o formato DD/MM/AAAA.")
        total_motoristas_label.config(text=f"Total de Motoristas: {total_motoristas}")
        return
    messagebox.showwarning("Erro", "Por favor, selecione motoristas ou especifique um período de data.")
    total_motoristas_label.config(text="Total de Motoristas: 0")
    messagebox.showwarning("Erro", "Por favor, selecione motoristas ou especifique um período de data.")

df_folgas = calcular_proxima_folga(df)

def get_executable_directory():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))
images_directory = get_executable_directory()

root = tk.Tk()
root.title("Consulta de Folgas de Motoristas")
root.geometry("1346x720")
root.configure(bg="#ffffff")
canvas = tk.Canvas(root, bg="#ffffff", height=720, width=1346, bd=0, highlightthickness=0, relief="ridge")
canvas.place(x=0, y=0)

img0_path = os.path.join(images_directory, 'img0.png')
img1_path = os.path.join(images_directory, 'img1.png')
img2_path = os.path.join(images_directory, 'background.png')
img3_path = os.path.join(images_directory, 'img_textBox0.png')
img4_path = os.path.join(images_directory, 'img_textBox1.png') 
img0 = Image.open(img0_path)
img0 = ImageTk.PhotoImage(img0)
img1 = Image.open(img1_path)
img1 = ImageTk.PhotoImage(img1)
img2 = Image.open(img2_path)
img2 = ImageTk.PhotoImage(img2)
img3 = Image.open(img3_path)
img3 = ImageTk.PhotoImage(img3)
img4 = Image.open(img4_path)
img4 = ImageTk.PhotoImage(img4)

background = canvas.create_image(673.0, 360.0, image=img2)
entry0_bg = canvas.create_image(1026.0, 233.5, image=img3)
entry1_bg = canvas.create_image(1027.5, 465.0, image=img4)
nome_var = tk.StringVar()
nomes_motoristas = df_folgas['Nome'].tolist()
nome_dropdown = ttk.Combobox(root, textvariable=nome_var, values=nomes_motoristas)
nome_dropdown.config(width=50)
nome_dropdown.bind("<Return>", lambda event: nome_var.set(', '.join(nome_dropdown.get().split(';'))))
canvas.create_window(1026.0, 233.5, window=nome_dropdown)

total_motoristas = len(nomes_motoristas)
total_motoristas_label = tk.Label(root, text=f"Total de Motoristas: {total_motoristas}", bg="#ffffff")
canvas.create_window(1208.0, 610.0, window=total_motoristas_label)

data_inicio_var = tk.StringVar()
data_inicio_label = tk.Label(root, text="Data Início (DD/MM/AAAA):", bg="#ffffff")
canvas.create_window(860.0, 30, window=data_inicio_label)
data_inicio_entry = ttk.Entry(root, textvariable=data_inicio_var)
canvas.create_window(1026.0, 30, window=data_inicio_entry)
data_fim_var = tk.StringVar()
data_fim_label = tk.Label(root, text="Data Fim (DD/MM/AAAA):", bg="#ffffff")
canvas.create_window(860, 60, window=data_fim_label)
data_fim_entry = ttk.Entry(root, textvariable=data_fim_var)
canvas.create_window(1026.0, 60, window=data_fim_entry)
resultado_textbox = tk.Text(root, width=60, height=15)
canvas.create_window(1026.0, 465.0, window=resultado_textbox)

img0 = tk.PhotoImage(file=img0_path)
b0 = tk.Button(image=img0, borderwidth=0, highlightthickness=0, command=consultar_motoristas, relief="flat")
b0.place(x=949, y=630, width=153, height=42)

img1 = tk.PhotoImage(file=img1_path)
b1 = tk.Button(image=img1, borderwidth=0, highlightthickness=0, command=exportar_excel, relief="flat")
b1.place(x=1192, y=672, width=149, height=32)
root.mainloop()
