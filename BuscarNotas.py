import pyautogui as py
import time as tm
import pandas as pd
from copiarcolarProdutos import copiarNfs,copiarProduts
from acoesBotao import bnt_selecionar_campos,check_desmacar_todos,check_empresa,check_num_nf,btn_ok,emitir,btn_abrir_origem,nota_encontrada,btn_fechar_aba,entrada_pesquisa_relatorio

import pyperclip

"""""
Passo 5 - ok

É necessário ter esc em toda abertura de abas para evitar possíves notificações

"""


def pesquisaNotas(numNotas):
    

    tempo = 10
   
    ##clica no botão pesquisar 
    # tm.sleep(tempo)
    py.click(x=375, y=33,clicks=2,interval=1)
   

    #faz a pesquisa pelo campo da nota fiscal
    # py.write('Nota Fiscal Emiss')
    pyperclip.copy('Nota Fiscal Emiss') 
    valor_copiado = pyperclip.paste()
    
    py.write(valor_copiado)
    tm.sleep(tempo)
    py.press('enter')
    
    # #Pesquisa a nota
    tm.sleep(10)
    py.press('esc', presses=2)
    tm.sleep(10)
    bnt_selecionar_campos()
    check_desmacar_todos()
    check_empresa()
    check_num_nf()
    btn_ok()
    

    emitir()
    
 
    
    df = pd.read_excel(numNotas)
    
    nf = df['N. Fiscal']
    # loop para percorrer linha por linha  das notas
    for nota in nf:
        #clica barra de pesquisa do numero
        py.click(x=187, y=119)
        tm.sleep(5)
        py.hotkey('ctrl','A')
        tm.sleep(5)
        py.press('delete')
        tm.sleep(5)

        #Abre o filtro avançado e pesquisa pelo numero e empresa
        py.write(str(nota))
        copiarNfs()
        tm.sleep(5)
        py.press('enter')
        tm.sleep(5)
        py.write('LOKAN')
        tm.sleep(5)
        py.press('enter',presses=2)
        tm.sleep(5)

        nota_encontrada()

    
        btn_abrir_origem()
        


        #clique aba produtos - esc para fechar notificações- usar tempo alto 15 - 30 segundos
        tm.sleep(15)
        py.press('esc',presses=2)
        py.hotkey('alt','4')
        tm.sleep(tempo)
        py.press('esc',presses=2)
   

     
        tm.sleep(10)
        
        copiarProduts()
        tm.sleep(0.5)
        py.press('esc')

     
        tm.sleep(4)
        py.press('esc',presses=2)
        btn_fechar_aba()
        tm.sleep(4)
        py.press('esc',presses=2)
        btn_fechar_aba()
     


