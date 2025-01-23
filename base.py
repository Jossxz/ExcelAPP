# base v0.3

# O import Abaixo, serve para o acesso da lista de Copiar e colar da Máquina 
import pyperclip
# O Import Abaixo, serve para indentificar Inputs pressionados no Teclado
import keyboard
# O import Abaixo, Serve para Executar comandos diretos no Excel utilizando de um webHook
from openpyxl import load_workbook

env = '16.01.0001.xlsx'  # Substitua com o caminho do seu arquivo
wb = load_workbook(env)
sheet = wb.active

sheet ['A2'] = '666.666.666-12'
sheet ['B2'] = 'Roberto Carlos Da Silva'
sheet ['C2'] = '12121'
sheet ['D2'] = 'AG'
sheet ['E2'] = 'X'
sheet ['F2'] = 'Cancelamento'
sheet ['G2'] = 'Total'
sheet ['H2'] = 400
sheet ['I2'] = '7780'
wb.save(env)

# Colar o texto da área de transferência
#texto_copiado = pyperclip.paste()
#print(f"O texto colado é: {texto_copiado}")

# CONTROLES
#press = keyboard.KEY_DOWN = "down"
#print({press})