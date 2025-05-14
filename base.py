import pyperclip
import keyboard
import xlwings as xw
import time
import tkinter as tk
from tkinter import filedialog

# Cria uma janelinha oculta apenas para abrir o diálogo de arquivo
root = tk.Tk()
root.withdraw()  # Oculta a janela principal

# Abre o seletor de arquivos
excel_path = filedialog.askopenfilename(
    title="Selecione o arquivo Excel",
    filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
)

# Se nenhum arquivo for selecionado, encerra o programa
if not excel_path:
    print("Nenhum arquivo selecionado. Encerrando.")
    exit()

# Abre o Excel visível
wb = xw.Book(excel_path, visible=True)
sheet = wb.sheets[0]

print("Pressione 'ç' para colar o texto do clipboard na célula atual e mover para a direita.")
print("Pressione 'esc' para sair.")

try:
    while True:
        if keyboard.is_pressed('ç'):
            txt_clip = pyperclip.paste()

            # Célula selecionada no momento
            current_cell = sheet.api.Application.ActiveCell

            row = current_cell.Row
            col = current_cell.Column

            print(f"Colando '{txt_clip}' em {chr(64 + col)}{row}")
            sheet.cells(row, col).value = txt_clip

            # Move para a célula à direita
            sheet.cells(row, col + 1).select()

            time.sleep(0.3)  # Evita colagens repetidas rápidas

        elif keyboard.is_pressed('esc'):
            break

        time.sleep(0.05)
finally:
    wb.save()
    wb.close()
    print("Excel salvo e fechado.")
