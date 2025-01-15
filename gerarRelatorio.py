               
import pyautogui as py
import time as tm
from caminhos import CAMINHO_NOTAS_DEBITOS,CAMINHO_NOTAS_CANCELADAS,CAMINHO_NOTAS_CREDITOS,EMPRESAS,CAMINHO_NOTAS_PAGAMENTOS
from logger import infoLogs
from acoesBotao import icon_contas_receber,btn_relatorio_car,verifica_texto_pagamentos,verifica_btn_Ok,verifica_texto_titulos_relatorio_aberto,btn_relatorios,btn_notas_canceladas,btn_relatorio_cap,btn_icone_contas_a_pagar,verifica_texto_notas_creditos,verifica_texto_titulos_relatorio,verifica_texto_marcar_todos,relatorio_faturamento_geral,bnt_muda_empresa,btn_fechar_aba,icon_salvar,notas_fiscais,verifica_bem_vindo_login,verifica_salvar_relatorio,verifica_icone_excel_superior,verifica_icone_relatorio,verifica_texto_notas_canceladas,verifica_vazio_notas_debitos,verifica_img_lokan_relatorio
from datas import carregar_datas
import pyperclip 
import os
from configuracao_tela import verifica_tamanho_tela,clique_nota_debito
from login import  login, fechar_sisloc

"""
Passo 2 - ok

Gera o Relatório do dia anterior após o login

"""

empresa = ''
def fechar_relatorio():
    btn_fechar_aba()
    py.press('esc',presses=2,interval=1)
    tm.sleep(5)
    btn_fechar_aba()
    
    



# Exemplo de uso para acessar as datas

def gerarRelatorioNotasCanceladas(empresa):
    datas = carregar_datas()

    periodo = ""
    
    infoLogs().info(f"ETAPA GERAR RELATORIO NOTAS DEBITOS CANCELADAS {empresa} - INICIADA")
    
    tempo = 2.5
     
    py.press('esc')
    tm.sleep(3)


    infoLogs().info("Pesquisar Relatorio de NOTAS DEBITOS CANCELADAS")
    #Clicando no menu para pesquisar Relatorio de Faturamento Geral 
    # tm.sleep(10)
    btn_relatorios()
    
    btn_notas_canceladas()
    
    while True:
        icone_relatorio_disponivel = verifica_icone_relatorio()
        if 'localizado' in icone_relatorio_disponivel:
            break
        else:
            print(f"Ainda não localizado verifica_icone_relatorio, tentando novamente... | {empresa}")
    
    
    py.press('tab',presses=3)
 
    tm.sleep(1)

    num_debito = ''

    

    if 'lokan' == empresa:
        num_debito = '14'#nº referente a nota debito cancelada lokan
    else:
        num_debito = '30'#nº referente a nota debito cancelada lokan_facil
    

    infoLogs().info(f'empresa em  notas canceladas {empresa} - {num_debito}')

    py.write(num_debito)#nº referente a nota debito cancelada

    py.press('enter')
    tm.sleep(1.5)
    py.press('tab',presses=2)
    tm.sleep(0.5)
    
    if isinstance(datas, tuple):
        # Se forem duas datas, imprime as duas
        py.write(datas[0])
    elif datas:
        py.write(datas)

    
    
    tm.sleep(0.5)
    py.press('tab')
    
    if isinstance(datas, tuple):
        # Se forem duas datas, imprime as duas
        py.write(datas[1])
    elif datas:
        py.write(datas)
    

    
        # tm.sleep(0.5)
        # py.press('tab',presses=2)
    tm.sleep(0.5)
    py.press('enter',presses=5)
        
    if not 'localizado' in verifica_img_lokan_relatorio():
        verifica_img_lokan_relatorio()
    else:
        while True:
            relatori_notas_canceladas_vazio = verifica_texto_notas_canceladas()
            if 'sim' in relatori_notas_canceladas_vazio:
                print("Não tem notas canceladas para o periodo informado...")
                btn_fechar_aba()
                infoLogs().info("ETAPA GERAR RELATORIO NOTAS DEBITOS CANCELADAS - FINALIZADA")
                break
            elif 'nao' in relatori_notas_canceladas_vazio:
                print("Tem notas canceladas...")
                
                icon_salvar()
                tm.sleep(tempo)
                
                # Criar o nome do arquivo baseado nas datas
                if isinstance(datas, tuple):
                    # Se forem duas datas, cria o período
                    periodo = rf'{datas[0]}_{datas[1]}'
                    
                    # Define o nome do arquivo com o período
                    nome_arquivo = rf'notas_canceladas_{periodo}_{empresa}.xls'
                else:
                    # Caso seja uma única data
                    nome_arquivo = rf'notas_canceladas_{datas}_{empresa}.xls'

                # Criar o caminho completo para o arquivo
                caminho_completo = os.path.join(CAMINHO_NOTAS_CANCELADAS, nome_arquivo)

                # Copiar o caminho completo para a área de transferência e colar (opcional)
                pyperclip.copy(caminho_completo)
        
                py.write(pyperclip.paste())  # Cola o texto final
        
               
                
                tm.sleep(1.5)
                py.press('enter')
                
                
                while True:
                    resultado = verifica_icone_excel_superior()  # Armazena o resultado da função
                    if 'localizado' in resultado:
                        break  # Sai do loop se 'localizado' for encontrado
                    else:
                        print("Ainda não localizado icone_excel_superior, tentando novamente...")
                
                tm.sleep(tempo)
                py.keyDown('alt')
                tm.sleep(tempo)
                py.press('f4')
                tm.sleep(tempo)
                py.keyUp('alt')
                tm.sleep(tempo)
                btn_fechar_aba()
                infoLogs().info(f"ETAPA GERAR RELATORIO NOTAS DEBITOS CANCELADAS  {empresa} - FINALIZADA")
                
              
                
                break






            


def gerarRelatorioCAP(empresa):
    datas = carregar_datas()

    periodo = ""
    
    infoLogs().info(f"ETAPA GERAR RELATORIO CAP LOKAN  {empresa} - INICIADA")
    
    tempo = 2.5
     
    
     
    py.press('esc')
    tm.sleep(3)
    
    #Clicando no menu para pesquisar Relatorio de Faturamento Geral 
    # tm.sleep(10)
    
    btn_icone_contas_a_pagar()
    
    btn_relatorios()
    
    btn_relatorio_cap()
    
    
    while True:
        icone_relatorio_disponivel = verifica_icone_relatorio()
        if 'localizado' in icone_relatorio_disponivel:
            break
        else:
            print("Ainda não localizado verifica_icone_relatorio, tentando novamente...")
            
    
    verifica_texto_titulos_relatorio()
    # py.press('tab')
    # # verifica_texto_marcar_todos()
    # py.press('tab',presses=18)
    # # verifica_texto_marcar_todos()
    py.press('tab',presses=15) #liquidação 1º data
    
    if isinstance(datas, tuple):
        # Se forem duas datas, imprime as duas
        py.write(datas[0])
    elif datas:
        py.write(datas)

    
    
    tm.sleep(0.5)
    py.press('tab')
    
    if isinstance(datas, tuple):
        # Se forem duas datas, imprime as duas
        py.write(datas[1])
    elif datas:
        py.write(datas)
        
    py.press('tab',presses=4)
    py.press('enter',presses=3)
    
    
    
    if not 'localizado' in verifica_img_lokan_relatorio():
        verifica_img_lokan_relatorio()
    else:
        while True:
            relatori_notas_creditos_vazio = verifica_texto_notas_creditos()
            if 'sim' in relatori_notas_creditos_vazio:
                print(f"Não tem notas creditos para o periodo informado... | {empresa}")
                btn_fechar_aba()
                tm.sleep(1)
                py.press('esc')
                tm.sleep(1)
                btn_fechar_aba()
                tm.sleep(1)
                py.press('esc')
                infoLogs().info(f"ETAPA GERAR RELATORIO CAP LOKAN {empresa} - FINALIZADA")
                break
            elif 'nao' in relatori_notas_creditos_vazio:
                print("Tem notas de creditos...")
                
                icon_salvar()
                tm.sleep(tempo)
                
                # Criar o nome do arquivo baseado nas datas
                if isinstance(datas, tuple):
                    # Se forem duas datas, cria o período
                    periodo = f'{datas[0]}_{datas[1]}'
                    
                    # Define o nome do arquivo com o período
                    nome_arquivo = f'notas_creditos_{periodo}_{empresa}.xls'
                else:
                    # Caso seja uma única data
                    nome_arquivo = f'notas_creditos_{datas}_{empresa}.xls'

                # Criar o caminho completo para o arquivo
                caminho_completo = os.path.join(CAMINHO_NOTAS_CREDITOS, nome_arquivo)

                # Copiar o caminho completo para a área de transferência e colar (opcional)
                pyperclip.copy(caminho_completo)
        
                py.write(pyperclip.paste())  # Cola o texto final
                    
               
                
                tm.sleep(1.5)
                py.press('enter')
                
                
                while True:
                    resultado = verifica_icone_excel_superior()  # Armazena o resultado da função
                    if 'localizado' in resultado:
                        break  # Sai do loop se 'localizado' for encontrado
                    else:
                        print("Ainda não localizado icone_excel_superior, tentando novamente...")

                tm.sleep(tempo)
                py.keyDown('alt')
                tm.sleep(tempo)
                py.press('f4')
                tm.sleep(tempo)
                py.keyUp('alt')
                tm.sleep(tempo)
                btn_fechar_aba()
                
                infoLogs().info(f"ETAPA GERAR RELATORIO CAP LOKAN {empresa} - FINALIZADA")
                
           
                                
                break



def gerar_relatorio_cap_aberto(empresa):
    datas = carregar_datas()

    periodo = ""
    
    infoLogs().info(f"ETAPA GERAR RELATORIO CAP EM ABERTO LOKAN {empresa} - INICIADA")
    
    tempo = 2.5
     
    # verifica = verifica_bem_vindo_login()
    
    # if 'localizado' not in verifica:
        
    #     gerarRelatorioCAP()

    # else:
    #    pass

     
    py.press('esc')
    tm.sleep(3)
    
   
    
    btn_icone_contas_a_pagar()
    
    btn_relatorios()
    
    btn_relatorio_cap()
    
    
    while True:
        icone_relatorio_disponivel = verifica_icone_relatorio()
        if 'localizado' in icone_relatorio_disponivel:
            break
        else:
            print("Ainda não localizado verifica_icone_relatorio, tentando novamente...")
            
    
    verifica_texto_titulos_relatorio_aberto()
    # py.press('tab')
    # # verifica_texto_marcar_todos()
    # py.press('tab',presses=18)
    # # verifica_texto_marcar_todos()
    py.press('tab',presses=11) #emissão 1º data
    
    if isinstance(datas, tuple):
        # Se forem duas datas, imprime as duas
        py.write(datas[0])
    elif datas:
        py.write(datas)

    
    
    tm.sleep(0.5)
    py.press('tab')
    
    if isinstance(datas, tuple):
        # Se forem duas datas, imprime as duas
        py.write(datas[1])
    elif datas:
        py.write(datas)
        
    py.press('tab',presses=5)
    tm.sleep(0.5)
    py.press('enter',presses=5)
    
    
    
    if not 'localizado' in verifica_img_lokan_relatorio():
        verifica_img_lokan_relatorio()
    else:
        while True:
            relatori_notas_creditos_vazio = verifica_texto_notas_creditos()
            if 'sim' in relatori_notas_creditos_vazio:
                print(f"Não tem notas creditos para o periodo informado... | {empresa}")
                btn_fechar_aba()
                tm.sleep(1)
                py.press('esc')
                tm.sleep(1)
                btn_fechar_aba()
                tm.sleep(1)
                py.press('esc')
                infoLogs().info(f"ETAPA GERAR RELATORIO CAP LOKAN {empresa} - FINALIZADA")
                break
            elif 'nao' in relatori_notas_creditos_vazio:
                print("Tem notas de creditos...")
                
                icon_salvar()
                tm.sleep(tempo)
                
                # Criar o nome do arquivo baseado nas datas
                if isinstance(datas, tuple):
                    # Se forem duas datas, cria o período
                    periodo = f'{datas[0]}_{datas[1]}'
                    
                    # Define o nome do arquivo com o período
                    nome_arquivo = f'notas_creditos_{periodo}_aberto_{empresa}.xls'
                else:
                    # Caso seja uma única data
                    nome_arquivo = f'notas_creditos_{datas}_aberto_{empresa}.xls'

                # Criar o caminho completo para o arquivo
                caminho_completo = os.path.join(CAMINHO_NOTAS_CREDITOS, nome_arquivo)

                # Copiar o caminho completo para a área de transferência e colar (opcional)
                pyperclip.copy(caminho_completo)
        
                py.write(pyperclip.paste())  # Cola o texto final
                    
               
                
                tm.sleep(1.5)
                py.press('enter')
                
                
                while True:
                    resultado = verifica_icone_excel_superior()  # Armazena o resultado da função
                    if 'localizado' in resultado:
                        break  # Sai do loop se 'localizado' for encontrado
                    else:
                        print("Ainda não localizado icone_excel_superior, tentando novamente...")

                tm.sleep(tempo)
                py.keyDown('alt')
                tm.sleep(tempo)
                py.press('f4')
                tm.sleep(tempo)
                py.keyUp('alt')
                tm.sleep(tempo)
                btn_fechar_aba()
                
                infoLogs().info(f"ETAPA GERAR RELATORIO CAP LOKAN {empresa} - FINALIZADA")
                
            
                                
                break


           
def cliques_status_1920x1080():
    py.click(924,381)
    tm.sleep(0.5)
    py.click(924,398)
    tm.sleep(0.5)
    py.click(924,413)
    tm.sleep(0.5)
    py.click(1061,381)
    tm.sleep(0.5)

    #LIQUIDADO
    py.click(1061,396)
    tm.sleep(0.5)
    py.click(1054,415)



def cliques_status_1366x768():
    py.click(x=646,y=277)
    tm.sleep(0.5)
    py.click(x=646,y=290)
    tm.sleep(0.5)
    py.click(x=646,y=307)
    tm.sleep(0.5)
    py.click(x=780,y=275)
    
    #Seleciona LIQUIDADO A COMPENSAR E LIQUIDADO
    tm.sleep(0.5)
    py.click(x=779,y=293)
    tm.sleep(0.5)
    py.click(x=774,y=304)


def gerarRelatorioCarPG(empresa):
    # Exemplo de uso para acessar as datas
    datas = carregar_datas()

    periodo = ""
    
    infoLogs().info(f"ETAPA GERAR RELATORIO CAR PG {empresa} - INICIADA")
    tempo = 1.5
     
    #Clicando no menu para pesquisar Relatorio de Faturamento Geral 
   # tm.sleep(10)
    
    # py.press('esc',presses=2,interval=2) #fechar aba
    
    icon_contas_receber()
    
    
    btn_relatorios()
 
    btn_relatorio_car()
   
            

    while True:
            icone_relatorio_disponivel = verifica_icone_relatorio()
            if 'localizado' in icone_relatorio_disponivel:
                break
            else:
                print("Ainda não localizado icone_relatorio, tentando novamente...")

    # #tira a seleção dos status
    # tm.sleep(10)
    
    tamanho_tela = verifica_tamanho_tela()

    if '1920x1080' in tamanho_tela:
        cliques_status_1920x1080()   

        # #CLICA EM CONTAS E MARCA TODOS
        py.press('tab',presses=5)
        tm.sleep(0.5)
        py.click(x=924,y=525)  
        #CLICA EM DATA LIQUIDAÇÃO 1ª CAMPO
        py.press('tab',presses=33)                      
        
    elif '1366x768' in tamanho_tela:
        cliques_status_1366x768()
        py.press('tab',presses=2)
    
             
        # #CLICA EM CONTAS E MARCA TODOS
        py.press('tab',presses=3)
        tm.sleep(1)
        py.click(x=680,y=414,clicks=2)
        #CLICA EM DATA LIQUIDAÇÃO 1ª CAMPO
        py.press('tab',presses=18)
    

    py.sleep(0.5)
    
    if isinstance(datas, tuple):
        # Se forem duas datas, imprime as duas
        py.write(datas[0])
    elif datas:
        py.write(datas)  
    
    tm.sleep(0.5)
    py.press('tab')
    
    # 2º DATA
    
    if isinstance(datas, tuple):
        # Se forem duas datas, imprime as duas
        py.write(datas[1])
    elif datas:
        py.write(datas) 
    
    tm.sleep(0.5)
    
    #GERAR RELATÓRIO
    py.press('tab',presses=7)
    py.press('enter',presses=5)
        
    if not 'localizado' in verifica_img_lokan_relatorio():
        tm.sleep(2)
        verifica_img_lokan_relatorio()
    else: 
        while True:
            verifica_tex_pagamentosCAR = verifica_texto_pagamentos()
            if 'localizado' in verifica_tex_pagamentosCAR:
                infoLogs().info(f"Não tem dados disponiveis para o periodo informado RELATORIO CAR | {empresa}")
                btn_fechar_aba()
                infoLogs().info(f"ETAPA GERAR RELATORIO CAR PG {empresa}- FINALIZADA")
                break
            elif 'seguir' in verifica_tex_pagamentosCAR:
                icon_salvar()
                
                while True:
                    resultado = verifica_salvar_relatorio()  # Armazena o resultado da função
                    if 'localizado' in resultado:
                        break  # Sai do loop se 'localizado' for encontrado
                    else:
                        icon_salvar()
                        print("Ainda não localizado - salvar_relatorio, tentando novamente...")
                        
             
                tm.sleep(tempo)
                
                if isinstance(datas, tuple):
                    # Se forem duas datas, imprime as duas
                    periodo = f'{datas[0]}_{datas[1]}'
                
                    nome_arquivo = f'pagamentos_{periodo}_{empresa}.xls'
                elif datas:
           
                    nome_arquivo = f'pagamentos_{datas}_{empresa}.xls'
                    
                    
                 # Criar o caminho completo para o arquivo
                caminho_completo = os.path.join(CAMINHO_NOTAS_PAGAMENTOS, nome_arquivo)

                # Copiar o caminho completo para a área de transferência e colar (opcional)
                pyperclip.copy(caminho_completo)
        
                py.write(pyperclip.paste())  # Cola o texto final
                
                
                
                tm.sleep(tempo)
                py.press('enter')
                
                
                while True:
                    resultado = verifica_icone_excel_superior()  # Armazena o resultado da função
                    if 'localizado' in resultado:
                        break  # Sai do loop se 'localizado' for encontrado
                    else:
                        print("Ainda não localizado icone_excel_superior, tentando novamente...")
                

                tm.sleep(tempo)
                py.hotkey('alt','F4',interval=0.5)
                # tm.sleep(tempo)
                # py.press('f4')
                # tm.sleep(tempo)
                # py.keyUp('alt')
                tm.sleep(tempo)

                verifica_bem_vindo_login()


                # py.press('esc',presses=2)
                btn_fechar_aba()
                infoLogs().info(f"ETAPA GERAR RELATORIO CAR PG {empresa}- FINALIZADA")
                
                break


def gerarRelatorio():
    datas = carregar_datas()

    periodo = ""
    
    infoLogs().info(f"ETAPA GERAR RELATORIO FATURAMENTO - INICIADA")
    
    tempo = 2.5

    verifica_btn_Ok()
    
    verifica = verifica_bem_vindo_login()
    
    if 'localizado' not in verifica:
        gerarRelatorio()

    else:
        ##mudar de empresa
        #Tempo para abrir o programa - dependendo da internet deve usar um tempo alto
        py.press('esc')
        tm.sleep(3)
        py.click(263,644)
        
        tm.sleep(tempo)

        for i in EMPRESAS:

            verifica_btn_mudar_empresa = bnt_muda_empresa()

            if 'nao' in verifica_btn_mudar_empresa:
                print('Reiniciando o Sisloc...')
                fechar_sisloc()
                login()

            else:
                pass  
            
            infoLogs().info(f'Numero da empresa consultada - {i}')
            py.press('down',presses=i) #-> antes  LOKAN
            tm.sleep(tempo)
            py.press('enter')
            
            tm.sleep(8)#tempo para carregar a empresa
            
            notas_fiscais()
            
            # entrada_pesquisa_relatorio()

            infoLogs().info("Pesquisar Relatorio de Faturamento Geral")
            #Clicando no menu para pesquisar Relatorio de Faturamento Geral 
            btn_relatorios()
            
            achou_btn_faturamento = relatorio_faturamento_geral()
            
            if achou_btn_faturamento == 'nao':
                btn_relatorios() #clica no btn relatorios novamente
            else:
                infoLogs().info("Filtro para buscar notas de debito")
                
                py.press('tab',presses=2)

                tamanho_tela = verifica_tamanho_tela()

                if '1920x1080' in tamanho_tela:
                    py.click(989,466)
                    tm.sleep(0.5)                           
                    
                elif '1366x768' in tamanho_tela:
                    py.click(709,355)
                    tm.sleep(0.5)

        
                py.click(x=1100, y=238)

            
                tm.sleep(tempo)
                py.press('tab')

                #preenche campo de data
                tm.sleep(tempo)
                if isinstance(datas, tuple):
                    # Se forem duas datas, imprime as duas
                    py.write(datas[0])
                elif datas:
                    py.write(datas)
                

                #Seleciona o 2ª campo de Data
                tm.sleep(tempo)
                py.press('tab')

                #preenche campo da 2ª data
                tm.sleep(tempo)
                
                if isinstance(datas, tuple):
                    # Se forem duas datas, imprime as duas
                    py.write(datas[1])
                elif datas:
                    py.write(datas)


                #Aperta 5x enter para gerar o relatorio
                tm.sleep(tempo)
                py.press('enter',presses=5)
                infoLogs().info("Gerando o relatorio Faturamento Geral")

                if not 'localizado' in verifica_img_lokan_relatorio():
                    tm.sleep(2)
                    verifica_img_lokan_relatorio()
                else: 

                    if i == 1:
                        empresa = 'lokan'
                    else:
                        empresa = 'lokanfacil'  

                    tm.sleep(3)
                    while True:
                        tem_notas_debitos = verifica_vazio_notas_debitos()
                        if 'sim' in tem_notas_debitos:
                            print(f"Não tem notas debitos para o periodo informado... | {empresa}")
                            btn_fechar_aba()
                            
                            infoLogs().info(f"ETAPA GERAR RELATORIO FATURAMENTO {empresa}- FINALIZADA")
                            
                            

                            gerarRelatorioNotasCanceladas(empresa) #notas canceladas
                            
                            fechar_relatorio()
                            
                            gerarRelatorioCAP(empresa)

                            gerar_relatorio_cap_aberto(empresa)
                            
                            fechar_relatorio()
                            
                            gerarRelatorioCarPG(empresa)
                            
                            fechar_relatorio()
                            
                            py.hotkey('alt','home') #clica no botão iniciar e deixa as abas abertas
                            
                            break
                        elif 'nao' in tem_notas_debitos:
                            
                            icon_salvar()
                            
                            while True:
                                resultado = verifica_salvar_relatorio()  # Armazena o resultado da função
                                if 'localizado' in resultado:
                                    break  # Sai do loop se 'localizado' for encontrado
                                else:
                                    print("Ainda não localizado - salvar_relatorio, tentando novamente...")
                                    
                            
                            
                                    # Criar o nome do arquivo baseado nas datas
                            if isinstance(datas, tuple):
                                # Se forem duas datas, cria o período
                                periodo = f'{datas[0]}_{datas[1]}'
                                
                                # Define o nome do arquivo com o período
                                nome_arquivo = f'notas_debitos_{periodo}_{empresa}.xls'
                            else:
                                # Caso seja uma única data
                                nome_arquivo = f'notas_debitos_{datas}_{empresa}.xls'

                            # Criar o caminho completo para o arquivo
                            caminho_completo = os.path.join(CAMINHO_NOTAS_DEBITOS, nome_arquivo)

                            # Copiar o caminho completo para a área de transferência e colar (opcional)
                            pyperclip.copy(caminho_completo)
                    
                            py.write(pyperclip.paste())  # Cola o texto final
                            
                            
                            tm.sleep(tempo)
                            py.press('enter')
                            
                            
                            while True:
                                resultado = verifica_icone_excel_superior()  # Armazena o resultado da função
                                if 'localizado' in resultado:
                                    break  # Sai do loop se 'localizado' for encontrado
                                else:
                                    print("Ainda não localizado icone_excel_superior, tentando novamente...")
                                    

                            tm.sleep(tempo)
                            py.keyDown('alt')
                            tm.sleep(tempo)
                            py.press('f4')
                            tm.sleep(tempo)
                            py.keyUp('alt')
                            tm.sleep(tempo)
                
                            btn_fechar_aba()

                            infoLogs().info(f"ETAPA GERAR RELATORIO FATURAMENTO {empresa} - FINALIZADA")
                            
                            
                            infoLogs().info(f"ETAPA FORMATAR PLANILHA NOTAS DEBITOS {empresa}- INICIADA")
                            
                            fechar_relatorio()
                            
                        
                            gerarRelatorioNotasCanceladas(empresa) #notas canceladas
                            
                            fechar_relatorio() #clica no botão iniciar e deixa as abas abertas
                            
                            gerarRelatorioCAP(empresa)
                            
                            gerar_relatorio_cap_aberto(empresa)
                            
                            fechar_relatorio()
                            
                            gerarRelatorioCarPG(empresa)
                            
                            fechar_relatorio()
                            
                            py.hotkey('alt','home') #clica no botão iniciar e deixa as abas abertas
                            
                            break






