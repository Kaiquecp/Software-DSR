import subprocess
import pandas as pd
import os
import pyodbc


maxtrack_path = r'C:\Users\user\Documents\codigos\Automacao_KM_Maxtrack\MaxTrack.py'
print("Executando MaxTrack.py...")
try:
    subprocess.run(['python', maxtrack_path], check=True)
    print("MaxTrack.py executado com sucesso!")
except subprocess.CalledProcessError as e:
    print(f"Erro ao executar MaxTrack.py: {e}")
    exit()
folder_path = r'C:\Users\user\Downloads\KM'

csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]
if len(csv_files) != 1:
    print("Erro: A pasta deve conter exatamente um arquivo CSV.")
else:
    file_path = os.path.join(folder_path, csv_files[0])

    try:
        df = pd.read_csv(file_path, sep=';', on_bad_lines='skip')

        if 'Identificador/Placa' in df.columns and 'Distância (Km)' in df.columns:
            df = df[['Identificador/Placa', 'Distância (Km)']]
            df['Identificador/Placa'] = df['Identificador/Placa'].str.replace(r'^(\w{3})(\w+)', r'\1-\2', regex=True)
            df['Distância (Km)'] = (
                df['Distância (Km)']
                .astype(str) 
                .str.replace(',', '.', regex=False)
                .str.strip()
            )
            df['Distância (Km)'] = pd.to_numeric(df['Distância (Km)'], errors='coerce')
            df = df[df['Distância (Km)'].notna() & (df['Distância (Km)'] >= 0)]
            df_sum = df.groupby('Identificador/Placa', as_index=False)['Distância (Km)'].sum()
            df_sum['Distância (Km)'] = df_sum['Distância (Km)'].round(2)
            df_sum = df_sum.sort_values(by='Distância (Km)', ascending=True)
            
            conn = pyodbc.connect(
                f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER=SERVER;DATABASE=DATABASE;UID=user;PWD=PASSWORD"
            )
            query = """
            QUERY
            """
            sql_df = pd.read_sql(query, conn)
            conn.close()

            sql_df.columns = ['Identificador/Placa']
            placas_faltantes = sql_df[~sql_df['Identificador/Placa'].isin(df_sum['Identificador/Placa'])]

            if not placas_faltantes.empty:
                placas_faltantes['Distância (Km)'] = 0
                df_sum = pd.concat([df_sum, placas_faltantes], ignore_index=True)

            output_path = r'C:\Users\user\Downloads\resultado_placas.xlsx'
            df_sum.to_excel(output_path, index=False)

            print(f'Arquivo processado e salvo em: {output_path}')
        else:
            print('Erro: As colunas "Identificador/Placa" e/ou "Distância (Km)" não foram encontradas no arquivo CSV.')
    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")
    finally:
        os.remove(file_path)
        print(f"Arquivo {file_path} foi apagado.")
