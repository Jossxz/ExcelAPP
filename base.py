import pyperclip
import keyboard
import xlwings as xw
import time

excel_path = R"C:\Users\olive\Downloads\Projects\Projetos\Python\Excel_Config\16.01.0001.xlsx"

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
