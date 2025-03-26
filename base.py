import os
import pyperclip
import keyboard
from openpyxl import load_workbook
import win32com.client as win32

# Caminho para o arquivo Excel
link_ = R"C:\Users\olive\Downloads\Projects\Projetos\Python\Excel_Config\16.01.0001.xlsx"

# Inicializando o Excel e abrindo o arquivo
excel = win32.Dispatch('Excel.Application')
excel.Visible = True
workbook = excel.Workbooks.Open(link_)
sheet = workbook.ActiveSheet

# Função que imprime o conteúdo da área de transferência e coloca no Excel
def printClip():
    txt_clip = pyperclip.paste()  # Obtém o texto da área de transferência
    print(f"Texto copiado: {txt_clip}")
    sheet.Cells(2, 3).Value = txt_clip  # Coloca o texto na célula A2

# Associa o hotkey (atalho de teclado) para chamar a função
keyboard.add_hotkey('ç', printClip())

# Mantém o script rodando, aguardando a tecla ser pressionada
print("Pressione 'ç' para copiar o texto da área de transferência para o Excel.")
keyboard.wait('esc')  # O script aguarda até que a tecla 'esc' seja pressionada para sair

# Libera todos os hooks de teclado ao final
keyboard.unhook_all()
