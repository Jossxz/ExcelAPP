import os
# O import Abaixo, serve para o acesso da lista de Copiar e colar da Máquina 
import pyperclip
# O Import Abaixo, serve para indentificar Inputs pressionados no Teclado
import keyboard
# O import Abaixo, Serve para Executar comandos diretos no Excel utilizando de um webHook
from openpyxl import load_workbook
# interagir diretamente com o Excel aberto, manipulando as células e salvando os dados em tempo real
import win32com.client as win32

# Caminho absoluto para o arquivo
caminho_arquivo = R"C:\Users\olive\Downloads\Projects\Projetos\Python\Excel_Config\16.01.0001.xlsx"

# Verifique se o arquivo realmente existe no caminho
if os.path.exists(caminho_arquivo):
    print(f"Arquivo encontrado: {caminho_arquivo}")
else:
    print(f"Arquivo não encontrado: {caminho_arquivo}")

# Inicialize o Excel
excel = win32.Dispatch('Excel.Application')
excel.Visible = True  # Tornar o Excel visível

# Abra o arquivo
workbook = excel.Workbooks.Open(caminho_arquivo)

# Acesse a planilha ativa
sheet = workbook.ActiveSheet

# Adicionar dados nas células
sheet.Cells(1, 1).Value = "Nome"
sheet.Cells(1, 2).Value = "Idade"

# Adicionar mais dados em células subsequentes
sheet.Cells(2, 1).Value = "João"
sheet.Cells(2, 2).Value = 30

sheet.Cells(3, 1).Value = "Maria"
sheet.Cells(3, 2).Value = 25

# Salvar o arquivo no mesmo local (ou em um novo caminho)
workbook.Save()  # Salva no mesmo arquivo, se preferir outro nome, substitua o caminho

# Fechar o Excel
# workbook.Close()
# excel.Quit()

# Defina a sequência de teclas que você deseja capturar
def imprimir_texto_copiado():
    texto_copiado = pyperclip.paste()  # Pega o texto da área de transferência
    print(f"Texto copiado: {texto_copiado}")

# Defina a sequência de teclas para disparar a ação
sequencia_teclas = 'ctrl+alt+p'  # Exemplo: pressionar Ctrl + Alt + P

# Aguardar até que a sequência de teclas seja pressionada e executar a função
keyboard.add_hotkey(sequencia_teclas, imprimir_texto_copiado)

# Manter o script rodando para escutar as teclas pressionadas
keyboard.wait()

# CONTROLES
#press = keyboard.KEY_DOWN = "down"
#print({press})