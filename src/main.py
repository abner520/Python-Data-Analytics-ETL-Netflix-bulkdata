import pandas as pd
import os
import glob

# caminho para ler os arquivos
folder_path = 'src\\data\\raw'

# Lista de todos os arquivos de excel
excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

if not excel_files:
    print("Nenhum Arquivo compativel encontrado")
else:

    # Data Frame = tabela na memoria para guardar os conteudos dos arquivos
    dfs = []  

    for excel_files in excel_files:

        try: 
             # Leio o arquivo de excel
            df_temp = pd.read_excel(excel_files)
            # Pegar o nome do arquivo
            file_name = os.path.basename(excel_files)

            df_temp['filename'] = file_name

            # Criamos uma nova colun chamada location
            if 'brasil' in file_name.lower():
                df_temp['location'] = 'br'
            elif 'france' in file_name.lower():
                df_temp['location'] = 'fr'
            elif 'italian' in file_name.lower():
                df_temp['location'] = 'it'

            # Criamos uma nova coluna chamada campaign
            df_temp['campaign'] = df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')

            # guardar dados tratados dentro de uma dataframe comum
            dfs.append(df_temp)

        except Exception as a:
    
            print(f"Erro ao abrir o arquivo {excel_files} : {e}")

if dfs:
    # Concatena todas as tabelas salvas no dfs em uma unica tabela
    result = pd.concat(dfs, ignore_index=True)

    # Caminho de Saida
    output_file = os.path.join('src', 'data', 'ready', 'clean.xlsx')
    # Configurei o motor de escrita
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    # leva os dados do resultado a serem escrito no motor de excel configurado
    result.to_excel(writer, index=False)
    # salva o arquivo de excel
    writer._save()

else:
    print('Nenhum dado para ser salv')