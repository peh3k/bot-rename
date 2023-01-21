import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog



def get_column_values_from_xlsx(file_path, sheet_name, column_name):
    # Carrega o arquivo xlsx usando o pandas
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    # Retorna os valores da coluna especificada em uma lista
    return list(df[column_name])


def rename_pdf_file(directory, old_name, new_name):
    old_file_path = os.path.join(directory, old_name + '.pdf')
    new_file_path = os.path.join(directory, new_name + '.pdf')
    os.rename(old_file_path, new_file_path)

def count_pdf_files(directory):
    pdf_files = [f for f in os.listdir(directory) if f.endswith('.pdf')]
    return len(pdf_files)




def on_start_button_clicked():
    alfabeto = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h']
    lista_nomes_tabela = get_column_values_from_xlsx("nomes.xlsx", "Planilha1", "nome")
    current_directory = os.path.dirname(__file__)

    for i in range(count_pdf_files(current_directory)):
        
        rename_pdf_file(current_directory, alfabeto[i], lista_nomes_tabela[i])


root = tk.Tk()
root.title("Arquivo XLSX")

xlsx_file = tk.StringVar()

xlsx_file_label = tk.Label(root, text="Arquivo XLSX:")
xlsx_file_label.grid(row=0, column=0, padx=5, pady=5)

xlsx_file_entry = tk.Entry(root, textvariable=xlsx_file)
xlsx_file_entry.grid(row=0, column=1, padx=5, pady=5)


start_button = tk.Button(root, text="Iniciar", command=on_start_button_clicked)
start_button.grid(row=1, column=1, padx=5, pady=5)


root.mainloop()
   

