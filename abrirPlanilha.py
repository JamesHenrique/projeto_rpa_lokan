import pyautogui as py
import time as tm
from datetime import timedelta, datetime
from datas import todasDatas


""""
Passo 4 - ok

Antes de consultar a primeira nota - abrir a planilha executar esse codigo

"""


def abrirExcel():
    tempo = 7

    caminho = f'Notas Debitos {todasDatas()}.xlsx'
    
    tm.sleep(tempo)
    py.press('winleft')
    tm.sleep(tempo)
    py.write(caminho)
    tm.sleep(tempo)
    py.press('enter')
    tm.sleep(tempo)
    py.press('right',presses=7)
    tm.sleep(tempo)
    py.press('enter')
    tm.sleep(tempo)
    py.hotkey('alt','tab')


    
