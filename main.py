import pandas as pd
import os
import time

excel_file_path = 'data_example.xlsx'
data = pd.read_excel(excel_file_path)
folder_path = 'result'

template_file_path = 'template.txt'

with open(template_file_path, 'r', encoding='utf-8') as file:
    template = file.read()

output_lines = []
if not os.path.exists(folder_path):
    os.makedirs(folder_path)
for index, row in data.iterrows():
    nama = row['nama']
    angka = row['num']
    # bagian = row['bagian']
    # if not os.path.exists(f'{folder_path}\\{bagian}'):
    #     os.makedirs(f'{folder_path}\\{bagian}')
    # output_file_path = os.path.join(f'{folder_path}\\{bagian}',f'{nama}_undangan.txt')
    if not os.path.exists(f'{folder_path}'):
        os.makedirs(f'{folder_path}')
    output_file_path = os.path.join(f'{folder_path}',f'{nama}_result.txt')
    # personalized_content = template.replace('nama_tamu', nama).replace('link_undangan', link)
    personalized_content = template.replace('nama', nama).replace('number', str(angka))
    output_lines.append(personalized_content)
    with open(output_file_path, 'w', encoding='utf-8') as file:
        file.writelines(output_lines)
    output_lines = []
    print(f"Data telah disimpan ke {output_file_path}")
    time.sleep(0.01)

