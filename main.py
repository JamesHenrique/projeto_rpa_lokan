from efetuando_login import fazer_login
from gerarRelatorio import gerarRelatorio
from criarPlanilhaPadrao import excluir_arquivos_xls
from formatarPlanilhaExcel import formatacao_planilha_final

import os
from logger import infoLogs
import time as tm
import pyautogui as py
import messagebox as ms
from datas import todasDatas,carregar_datas
from efetuando_login import fechar_sisloc



while True:
    print('-------------------------- RPA LOKAN V3 --------------------------\n')
    esc = input(f'Deseja formatar planilhas baixadas do periodo {carregar_datas()} ?\n\n"S" para sim | "N" para não..: ').strip().lower()

    if esc == 's':
        print(f'Formatando planilhas do periodo {carregar_datas()}')
        excluir_arquivos_xls()
        formatacao_planilha_final()
        ms.showinfo("FINALIZADO RPA","Formatação finalizada com sucesso")
        break
    
    if esc == 'n':
        print('Saindo da formatação...')
        tm.sleep(3)
        os.system('cls')

        todasDatas()

        fazer_login()

        infoLogs().info("ETAPA LOGIN FINALIZADA")

        tm.sleep(10)


        gerarRelatorio()


        py.press('esc',presses=10,interval=0.5)


        fechar_sisloc()



        formatacao_planilha_final()


        excluir_arquivos_xls()
        

        ms.showinfo("FINALIZADO RPA","Automação finalizada com sucesso")
        break
    else:
        print('Escolha uma das opções para seguir | "S" para sim ou "N" para nao\n')
        continue