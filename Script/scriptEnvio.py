import openpyxl
import pyperclip
import pyautogui as pg
from time import sleep

t = 2.4  # Tempo de pausa entre ações

# === CARREGAR PLANILHA === #
work = openpyxl.load_workbook('BD_Ass.xlsx')
plan = work['Assinaturas']

pg.alert('SCRIPT INICIANDO')
pg.click(563,745)

for line in plan.iter_rows(min_row=2):
    if all(cell.value is None for cell in line):
        pg.alert('Linha vazia na planilha! Encerrando o script de ENVIO DE ASSINATURA.')
        break
    
    nome = line[0].value
    email = line[9].value
    assunto = "Nova Assinatura de E-mail"
    caminho_pasta = fr"C:\Users\nicor\OneDrive\Desktop\Assinaturas Novas\{nome}"
    
    # === ABRINDO OUTLOOK === #
    pg.click(217,25) #espaço em branco outlook
    sleep(t)
    
    # === ABRINDO NOVO E-MAIL === #
    pg.press('alt')
    pg.press('c')
    pg.press('n')
    pg.press('n')
    sleep(5)
    
    # === TROCANDO PARA SUPORTE.TI === #
    pg.click(144,174) #clique em "DE"
    sleep(t)
    pg.press('down')
    pg.press('enter')
    sleep(t)
    
    # === COLANDO E-MAIL === #
    pg.click(199,210) #clique em "PARA"
    pyperclip.copy(email)
    pg.hotkey('ctrl', 'v')
    sleep(t)
    pg.press('enter')
    sleep(t)
    
    # ==== COLANDO ASSUNTO === #
    pg.click(220,333) #clique em "Assunto"
    pyperclip.copy(assunto)
    pg.hotkey('ctrl', 'v')
    sleep(t)
    pg.press('enter')
    sleep(t)
    
    # === COLANDO CORPO === #
    pg.press('alt')
    pg.press('m')
    pg.press('p')
    pg.press('s')
    pg.press('down')
    pg.press('enter')
    sleep(t)
    pg.press('backSpace')
    pg.press('backSpace')
    sleep(t)
    
    # === ABRINDO A PASTA DA ASSINATURA === #
    pg.press('alt')
    pg.press('m')
    pg.press('a')
    pg.press('x')
    pg.press('p')
    sleep(t)
    
    pyperclip.copy(caminho_pasta)
    sleep(t)
    
    pg.hotkey('ctrl', 'v')
    pg.press('enter')
    sleep(t)
    
    pg.click(562,272) #clique em um espaço vazio da pasta do colaborador
    sleep(t)
    
    # ==== COPIANDO OS ARQUIVOS DA PASTA === #
    pg.hotkey('ctrl', 'a')
    pg.hotkey('enter')
    pg.click(58,181) #clique me enviar
    sleep(t)
    
pg.alert('SCRIPT FINALIZADO!')