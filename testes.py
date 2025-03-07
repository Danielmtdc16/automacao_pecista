from datetime import time, timedelta
from time import sleep

import pyautogui as pg
hora_entrada = timedelta(hours=7, minutes=47)
hora_saida_almoco = timedelta(hours=0, minutes=00)
hora_volta_almoco = timedelta(hours=0,minutes=30)
hora_saida = timedelta(hours=14, minutes=23)

total = hora_saida - hora_entrada - (hora_volta_almoco - hora_saida_almoco)
saida = hora_entrada + timedelta(hours=6, minutes=0)


imagem = 'msg_confirm_nao.png'



while True:
    print(pg.position())
    sleep(1)