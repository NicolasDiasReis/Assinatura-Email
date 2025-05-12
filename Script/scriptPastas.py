import openpyxl
import pyperclip
import pyautogui as pg
from time import sleep
import os
import shutil

# === CONFIGURAÇÕES === #
caminho_base = fr"C:\Users\nicor\OneDrive\Desktop\Assinaturas Novas"
caminho_gif = fr"C:\Users\nicor\OneDrive\Desktop\Assinaturas Novas\GIF"
t = 1.5  # Tempo de pausa entre ações
t_ai = 1  # Tempo de pausa entre ações no adobe
first_export = True

# === CARREGAR PLANILHA === #
work = openpyxl.load_workbook('BD_Ass.xlsx')
plan = work['Assinaturas']

pg.alert('SCRIPT INICIANDO')

# Abrir o Illustrator (ajuste coordenadas conforme necessário)
pg.click(608, 742) #Adobe Illustrator
sleep(t)
pg.click(656,139) #Espaço em branco do Adobe
sleep(t)

# === LOOP PARA CADA COLABORADOR === #
for line in plan.iter_rows(min_row=2):
    if all(cell.value is None for cell in line):
        pg.alert('Linha vazia na planilha! Encerrando o script de CRIAÇÃO DE ASSINATURA.')
        break
    
    nome = line[0].value
    dept = line[1].value
    ddd = line[2].value
    telefone = line[5].value
    email = line[9].value
    assunto = "Nova Assinatura de E-mail"
    
    # === EDITAR NO ILLUSTRATOR === #
    pg.hotkey('t')
    pg.click(156,486) #clique no nome
    pg.hotkey('ctrl', 'a')
    pg.press('del')
    pyperclip.copy(nome)
    pg.hotkey('ctrl', 'v')
    sleep(t_ai)
    
    pg.click(146,539) #clique no departarmento
    pg.hotkey('ctrl', 'a')
    pg.press('del')
    pyperclip.copy(dept)
    pg.hotkey('ctrl', 'v')
    sleep(t_ai)
    
    pg.doubleClick(752,404) #clique no ddd
    pg.press('del')
    pyperclip.copy(ddd)
    pg.hotkey('ctrl', 'v')
    sleep(t_ai)
    
    pg.click(815,403) #clique no telefone
    pg.hotkey('ctrl', 'a')
    pg.press('del')
    pyperclip.copy(telefone)
    pg.hotkey('ctrl', 'v')
    sleep(t_ai)
    
    # === ABRIR JANELA DE EXPORTAÇÃO === #
    pg.hotkey('alt')
    pg.hotkey('f')
    pg.hotkey('e')
    pg.hotkey('down')
    pg.hotkey('enter')
    sleep(t)
    
    # === NAVEGAR ATÉ O CAMINHO BASE === #
    for _ in range(8):
        pg.press('tab')
    pg.press('enter')
    pyperclip.copy(caminho_base)
    pg.hotkey('ctrl', 'v')
    pg.press('enter')
    sleep(t)
    
    # === CRIAR NOVA PASTA COM O NOME DO COLABORADOR === #
    pg.hotkey('ctrl', 'shift', 'n')
    sleep(t)
    pyperclip.copy(nome)
    pg.hotkey('ctrl', 'v')
    pg.press('enter')
    pg.press('enter')
    sleep(t)
    
    # === EXPORTAR ARQUIVO === #
    if first_export:
        for _ in range(2):  # ou qualquer número de repetições
            pg.press('tab')
            first_export = False
    else:
        pg.press('tab')
    sleep(t)
    
    #===================================^^^^^^^^^^^^=====================================#
    pg.press('right')
    pg.press('space')
    pg.hotkey('ctrl', 'v')
    sleep(t)
    
    for _ in range(4):
        pg.press('tab')
    sleep(t)
    
    pg.press('enter') # Confirmar exportação, se necessário
    
    for _ in range(4):
        pg.press('tab')
    pg.press('enter')
    sleep(t)
    
    # === COPIAR GIFS PARA PASTA DO COLABORADOR === #
    destino = os.path.join(caminho_base, nome)
    if not os.path.exists(destino):
        os.makedirs(destino)
    for arquivo in os.listdir(caminho_gif):
        if arquivo.lower().endswith('.gif'):
            origem_arquivo = os.path.join(caminho_gif, arquivo)
            destino_arquivo = os.path.join(destino, arquivo)
            shutil.copy(origem_arquivo, destino_arquivo)