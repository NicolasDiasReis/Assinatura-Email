import openpyxl
import pyperclip
import pyautogui as pg
from time import sleep
import os
import shutil

# === CONFIGURAÇÕES === #
caminho_base = fr"C:\Users\nicor\OneDrive\Desktop\Assinaturas Novas"
caminho_gif = fr"C:\Users\nicor\OneDrive\Desktop\Assinaturas Novas\GIF"
t = 2.5  # Tempo de pausa entre ações

# === CARREGAR PLANILHA === #
work = openpyxl.load_workbook('BD_Ass.xlsx')
plan = work['Assinaturas']

pg.alert('SCRIPT INICIANDO')

# Abrir o Illustrator (ajuste coordenadas conforme necessário)
pg.click(608, 742)
sleep(t)
pg.click(476, 115)

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
    assunto = "Nova Assinatura de E-mail"#
    # === EDITAR NO ILLUSTRATOR === #
    sleep(1)
    pg.hotkey('t')
    pg.click(196, 428)
    pg.hotkey('ctrl', 'a')
    pg.press('del')
    pyperclip.copy(nome)
    pg.hotkey('ctrl', 'v')
    sleep(t)#
    pg.doubleClick(190, 481)
    pg.hotkey('ctrl', 'a')
    pg.press('del')
    pyperclip.copy(dept)
    pg.hotkey('ctrl', 'v')
    sleep(t)#
    pg.doubleClick(774, 348)
    pg.press('del')
    pyperclip.copy(ddd)
    pg.hotkey('ctrl', 'v')
    sleep(t)#
    pg.click(839, 347)
    pg.hotkey('ctrl', 'a')
    pg.press('del')
    pyperclip.copy(telefone)
    pg.hotkey('ctrl', 'v')
    sleep(t)#
    # === ABRIR JANELA DE EXPORTAÇÃO === #
    pg.hotkey('alt')
    pg.hotkey('f')
    pg.hotkey('e')
    pg.hotkey('down')
    pg.hotkey('enter')
    sleep(t)#
    # === NAVEGAR ATÉ O CAMINHO BASE === #
    for _ in range(8):
        pg.press('tab')
    pg.press('enter')
    pyperclip.copy(caminho_base)
    pg.hotkey('ctrl', 'v')
    pg.press('enter')
    sleep(t)#
    # === CRIAR NOVA PASTA COM O NOME DO COLABORADOR === #
    pg.hotkey('ctrl', 'shift', 'n')
    sleep(t)
    pyperclip.copy(nome)
    pg.hotkey('ctrl', 'v')
    pg.press('enter')
    pg.press('enter')
    sleep(t)#
    # === EXPORTAR ARQUIVO === #
    for _ in range(2):
        pg.press('tab')
    pg.press('right')
    pg.press('space')
    pg.hotkey('ctrl', 'v')
    for _ in range(4):
        pg.press('tab')
    pg.press('enter')
    sleep(t)
    pg.press('enter')  # Confirmar exportação, se necessário#
    # === COPIAR GIFS PARA PASTA DO COLABORADOR === #
    destino = os.path.join(caminho_base, nome)
    if not os.path.exists(destino):
        os.makedirs(destino)#
    for arquivo in os.listdir(caminho_gif):
        if arquivo.lower().endswith('.gif'):
            origem_arquivo = os.path.join(caminho_gif, arquivo)
            destino_arquivo = os.path.join(destino, arquivo)
            shutil.copy(origem_arquivo, destino_arquivo)