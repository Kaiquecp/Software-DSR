import os
import sys
import pandas as pd
import pyodbc
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk

# Verifica se está rodando como executável ou script
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS  # Caminho temporário do PyInstaller
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

# Configurações de conexão
server = 'hs100.scp.tec.br,41433'
database = 'ATSLog_Transparana'
username = 'atscbd_Transparana'
password = 'atscbd@@12345'

# String de conexão ODBC
conn_str = (
    'DRIVER={ODBC Driver 18 for SQL Server};'
    f'SERVER={server};'
    f'DATABASE={database};'
    f'UID={username};'
    f'PWD={password};'
    'TrustServerCertificate=yes;'
)

# Criando a conexão com o banco de dados usando pyodbc
conn = pyodbc.connect(conn_str)

# Consulta SQL
query = """
Select
     GETDATE()  as "Atualizado"
     , P.pes_nome             AS "Nome"
     , (select top 1 gru_nome from grupo G join Pessoa_Grupo pg on g.gru_Codigo = pg.gru_Codigo where PG.Pes_Codigo = P.Pes_Codigo) as 'Grupo'
     , Eve_nome                    AS "Evento"
     , pmE_valor              AS "Tempo Evento"
     , ( pmE_valor / 3600 )  AS "Tempo Horas"
     , ((pmE_valor / 3600) -  FLOOR(pmE_valor / 3600))  *60  as "Minutos"
     , PCM.PCM_Data                AS "Data Base"
     , CASE 
          WHEN pmE_valor > 39600 THEN 1 --Prodelog é 11hrs de jornada aderente
          ELSE 0    
     END AS "Acima"
     , CASE 
          WHEN pmE_valor > 39600 THEN (pmE_valor - 39600)/60
          ELSE 0    
     END AS "Minutos Acima"   
     , Ope.Ope_Descricao as 'Operação'     
from Pessoa_Mot_Mov_Evento    PEV
     join Evento                   E         on e.Eve_Codigo     = PEV.Eve_Codigo
     join Pessoa_Mot_Ciclo_Mov     PCM       on PCM.PCM_Codigo   = PEV.PCM_Codigo
     join Pessoa                   P         on P.Pes_Codigo     = PCM.Pes_Codigo
     left join Pessoa_Operacao POP WITH (NOLOCK) ON POP.pes_Codigo = P.pes_Codigo
     left join Operacao Ope WITH (NOLOCK) ON Ope.Ope_Codigo = POP.Ope_Codigo

where 
     PEV.Eve_Codigo = 55
     and PEV.PME_Inativo = 0
     and pst_codigo = 8
     and PCM_Data between getdate() - 150 and getdate()+1
		   and P.Pes_Codigo in 
		   (select Pes_Codigo_Relacionado from Pessoa_Relacao where Pes_Codigo_Principal in (select Pes_Codigo from pessoa where pes_codigo = 1 or pes_codigo in(
		  (select Pes_Codigo_Relacionado from Pessoa_Relacao where TPR_Codigo = 24 and Pes_Codigo_Principal = 1)) 
		  and PSt_Codigo = 1)  and TPR_Codigo = 5)

order by  PCM.PCM_Data DESC
"""

# Executando a consulta SQL
df = pd.read_sql(query, conn)

# Função para calcular as folgas
def calcular_folgas(datas):
    folgas = []
    dias_trabalhados = 0
    ultimo_dia = None

    for i in range(len(datas)):
        if ultimo_dia is not None and (datas[i] - ultimo_dia).days > 1:
            while (ultimo_dia + timedelta(days=1)) < datas[i]:
                ultimo_dia += timedelta(days=1)
                folgas.append(ultimo_dia.strftime("%Y-%m-%d"))

            dias_trabalhados = 0  # Reinicia a contagem de dias trabalhados

        ultimo_dia = datas[i]
        dias_trabalhados += 1

        if dias_trabalhados == 6:
            folga = ultimo_dia + timedelta(days=1)
            folgas.append(folga.strftime("%Y-%m-%d"))
            dias_trabalhados = 0  # Reinicia a contagem de dias trabalhados

    if dias_trabalhados > 0 and dias_trabalhados < 6:
        dias_faltantes = 6 - dias_trabalhados
        folga = ultimo_dia + timedelta(days=dias_faltantes) + timedelta(days=1)
        folgas.append(folga.strftime("%Y-%m-%d"))

    return list(set(folgas))

# Função para calcular a próxima folga a partir do DataFrame
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

# Função para exportar o resultado para Excel
def exportar_excel():
    folgas_df = calcular_proxima_folga(df)  # Calcula as folgas antes de exportar
    arquivo_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')])
    if arquivo_path:
        folgas_df.to_excel(arquivo_path, index=False)
        messagebox.showinfo("Exportação", "Arquivo exportado com sucesso!")

def consultar_motoristas():
    motoristas_selecionados = nome_var.get().strip()
    data_inicio = data_inicio_var.get().strip()
    data_fim = data_fim_var.get().strip()

    resultado_textbox.delete('1.0', tk.END)  # Limpa o resultado anterior
    total_motoristas = 0  # Variável para contar motoristas

    # Consulta por motoristas
    if motoristas_selecionados:
        motoristas_lista = motoristas_selecionados.split(', ')
        folga_motorista = df_folgas[df_folgas['Nome'].isin(motoristas_lista)]
        total_motoristas = len(folga_motorista)  # Conta os motoristas
        if not folga_motorista.empty:
            resultado = ""
            for index, row in enumerate(folga_motorista.iterrows(), 1):  # Adiciona numeração
                _, data = row
                resultado += f"{index}. {data['Nome']},\n PRÓXIMA FOLGA: {data['Próxima Folga']}\n\n"  # Adiciona uma linha em branco
            resultado_textbox.insert(tk.END, resultado)
        else:
            resultado_textbox.insert(tk.END, "Motoristas não encontrados ou sem folgas calculadas\n")
        total_motoristas_label.config(text=f"Total de Motoristas: {total_motoristas}")  # Atualiza o total
        return

    # Consulta por período de data
    if data_inicio and data_fim:
        try:
            data_inicio_dt = datetime.strptime(data_inicio, "%d/%m/%Y")
            data_fim_dt = datetime.strptime(data_fim, "%d/%m/%Y")
            folgas_periodo = df_folgas[
                (pd.to_datetime(df_folgas['Próxima Folga'], format="%d/%m/%Y") >= data_inicio_dt) &
                (pd.to_datetime(df_folgas['Próxima Folga'], format="%d/%m/%Y") <= data_fim_dt)
            ]
            total_motoristas = len(folgas_periodo)  # Conta os motoristas
            if not folgas_periodo.empty:
                resultado = ""
                for index, row in enumerate(folgas_periodo.iterrows(), 1):  # Adiciona numeração
                    _, data = row
                    resultado += f"{index}. {data['Nome']},\n PRÓXIMA FOLGA: {data['Próxima Folga']}\n\n"  # Adiciona uma linha em branco
                resultado_textbox.insert(tk.END, resultado)
            else:
                resultado_textbox.insert(tk.END, "Nenhuma folga encontrada no período especificado\n")
        except ValueError:
            messagebox.showerror("Erro", "Formato de data inválido. Por favor, use o formato DD/MM/AAAA.")
        total_motoristas_label.config(text=f"Total de Motoristas: {total_motoristas}")  # Atualiza o total
        return

    # Caso nenhum nome ou data seja informado
    messagebox.showwarning("Erro", "Por favor, selecione motoristas ou especifique um período de data.")
    total_motoristas_label.config(text="Total de Motoristas: 0")  # Caso não haja motoristas

    # Caso nenhum nome ou data seja informado
    messagebox.showwarning("Erro", "Por favor, selecione motoristas ou especifique um período de data.")


# Calcular as próximas folgas
df_folgas = calcular_proxima_folga(df)

# Função para obter o diretório do executável
def get_executable_directory():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

# Configura o diretório onde as imagens estão localizadas
images_directory = get_executable_directory()

# Criando a interface gráfica
root = tk.Tk()
root.title("Consulta de Folgas de Motoristas")
root.geometry("1346x720")
root.configure(bg="#ffffff")
canvas = tk.Canvas(root, bg="#ffffff", height=720, width=1346, bd=0, highlightthickness=0, relief="ridge")
canvas.place(x=0, y=0)

#Carregando imagens
img0_path = os.path.join(images_directory, 'img0.png')  # Substitua pelo nome da sua imagem
img1_path = os.path.join(images_directory, 'img1.png')  # Substitua pelo nome da sua imagem
img2_path = os.path.join(images_directory, 'background.png')  # Substitua pelo nome da sua imagem
img3_path = os.path.join(images_directory, 'img_textBox0.png')  # Substitua pelo nome da sua imagem
img4_path = os.path.join(images_directory, 'img_textBox1.png')  # Substitua pelo nome da sua imagem
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
 

# Carregando as imagens

background = canvas.create_image(673.0, 360.0, image=img2)


entry0_bg = canvas.create_image(1026.0, 233.5, image=img3)


entry1_bg = canvas.create_image(1027.5, 465.0, image=img4)

# Lista suspensa de motoristas com seleção múltipla
nome_var = tk.StringVar()
nomes_motoristas = df_folgas['Nome'].tolist()  # Extrai os nomes de df_folgas
nome_dropdown = ttk.Combobox(root, textvariable=nome_var, values=nomes_motoristas)
nome_dropdown.config(width=50)
nome_dropdown.bind("<Return>", lambda event: nome_var.set(', '.join(nome_dropdown.get().split(';'))))

# Adicionando a Combobox ao Canvas
canvas.create_window(1026.0, 233.5, window=nome_dropdown)

# Contagem correta de motoristas
total_motoristas = len(nomes_motoristas)  # Corrige a contagem de motoristas
total_motoristas_label = tk.Label(root, text=f"Total de Motoristas: {total_motoristas}", bg="#ffffff")
canvas.create_window(1208.0, 610.0, window=total_motoristas_label)

# Entrada de data de início
data_inicio_var = tk.StringVar()
data_inicio_label = tk.Label(root, text="Data Início (DD/MM/AAAA):", bg="#ffffff")
canvas.create_window(860.0, 30, window=data_inicio_label)
data_inicio_entry = ttk.Entry(root, textvariable=data_inicio_var)
canvas.create_window(1026.0, 30, window=data_inicio_entry)

# Entrada de data de fim
data_fim_var = tk.StringVar()
data_fim_label = tk.Label(root, text="Data Fim (DD/MM/AAAA):", bg="#ffffff")
canvas.create_window(860, 60, window=data_fim_label)
data_fim_entry = ttk.Entry(root, textvariable=data_fim_var)
canvas.create_window(1026.0, 60, window=data_fim_entry)

# TextBox para exibir os resultados
resultado_textbox = tk.Text(root, width=60, height=15)
canvas.create_window(1026.0, 465.0, window=resultado_textbox)

# Botão para consultar motoristas
img0 = tk.PhotoImage(file=img0_path)
b0 = tk.Button(image=img0, borderwidth=0, highlightthickness=0, command=consultar_motoristas, relief="flat")
b0.place(x=949, y=630, width=153, height=42)

# Botão para exportar o DataFrame com as folgas para Excel
img1 = tk.PhotoImage(file=img1_path)
b1 = tk.Button(image=img1, borderwidth=0, highlightthickness=0, command=exportar_excel, relief="flat")
b1.place(x=1192, y=672, width=149, height=32)

# Iniciando o loop principal da interface gráfica
root.mainloop()