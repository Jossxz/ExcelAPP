# O import Abaixo, serve para o acesso da lista de Copiar e colar da Máquina 
import pyperclip
# O Import Abaixo, serve para indentificar Inputs pressionados no Teclado
import keyboard
# O import Abaixo, Serve para Executar comandos diretos no Excel utilizando de um webHook
from openpyxl import load_workbook

# interagir diretamente com o Excel aberto, manipulando as células e salvando os dados em tempo real
import win32com.client as win32

excel = win32.Dispatch('Excel.Application')
excel.Visible = True

workbook = excel.Workbooks.Open("16.01.0001.xlsx")

wb = load_workbook(env)
sheet = wb.active

sheet.Cells(1, 1).Value = "Nome"
sheet.Cells(1, 2).Value = "Idade"

# Adicionar mais dados em células subsequentes
sheet.Cells(2, 1).Value = "João"
sheet.Cells(2, 2).Value = 30

sheet.Cells(3, 1).Value = "Maria"
sheet.Cells(3, 2).Value = 25


wb.save(env)

# Colar o texto da área de transferência
#texto_copiado = pyperclip.paste()
#print(f"O texto colado é: {texto_copiado}")

# CONTROLES
#press = keyboard.KEY_DOWN = "down"
#print({press})