import datetime
from fileinput import close
from time import sleep

import pyautogui as pg
import pandas as pd
import openpyxl
from pyscreeze import ImageNotFoundException

# LOGIN PEDRO
user = '4150'
password = 'teles'

notas_fiscais = pd.read_excel('notas_fiscais_pecista.xlsx')
notas_fiscais['Vencimento'] = pd.to_datetime(notas_fiscais['Vencimento'], errors='coerce')

dia_atual = datetime.datetime.now().day
primeira_semana = dia_atual + 1 + 7
segunda_semana = primeira_semana + 7

notas_prorrogar = []

for index, nota in notas_fiscais.iterrows():
    if (dia_atual <= nota['Vencimento'].day <= primeira_semana) or (nota['Vencimento'].day > primeira_semana and nota['Status'] == 'Prorroga'):
        notas_prorrogar.append(
            {
                'cod_fornecedor': nota['Codfor'],
                'num_nota': nota['documento'][3:]
            }
        )

# Funções
def enter():
    pg.press('enter')

def digitar(msg, intervalo):
    pg.write(msg, interval=intervalo)

def clicar(tecla):
    pg.press(tecla)

pg.FAILSAFE = False
# SIAC Pyautogui
largura, altura = pg.size()

canto_inferior_direito = (largura - 1, altura - 1)

pg.moveTo(canto_inferior_direito)

pg.click()

clicar('winleft')
sleep(2)
digitar('pecista siac', 0.25)
sleep(2)
enter()
sleep(2)
clicar('2')
sleep(1)
digitar(user+password, 0.25)
sleep(5)
enter()
sleep(3)
clicar('9')
sleep(2)

caminho_imagem_confirma = 'msg_confirma.png'

for nota in notas_prorrogar:
    print('se to aqui e pq consegui')
    sleep(1)
    digitar(str(nota['cod_fornecedor']), 0.25)
    sleep(0.5)
    digitar(str(nota['num_nota']), 0.25)
    sleep(0.5)
    enter()

    while True:
        for i in range(5):
            clicar('right')
            sleep(0.4)
        try:
            img1 = pg.locateOnScreen(caminho_imagem_confirma, confidence=0.9)
            if img1:
                print('achei a mensagem de confirmacao')
                clicar('s')
                break
        except pg.ImageNotFoundException:
            continue
