import pyautogui
import time
import pyautogui as py
import psutil


# time.sleep(3)
# screenshot = pyautogui.screenshot(r'C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\btn_relatorios.png', region=(394,206 ,120, 40))#x,y,largura,altura




"""



744,530 - 
icon_usuario_loginV3 - 469,368

209,587 - texto_disponivel_notas_creditos1920x1080

121,620 - texto_disponivel_notas_creditos

210,648 - texto_disponivel_pagamentos1920x1080

183,337 - texto_disponivel_notas_canceladas1920x1080

1237,469 - texto_vazio_notas_debitos1920x1080

200,614 - texto_disponivel_notas_creditos

914,432 - texto_titulos_relatorio

921,498 - marcar todos

notas_debitos_22-10-2024

734,550

363,177 icone_salvar_relatorio

16,14  icone_excel_superior

214,211 - img_lokan_relatorio

1237,469 - texto_vazio_relatorio


"""



def verifica_img_lokan_relatorio():
    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\img_lokan_relatorio.png", confidence=0.3) #alterar o confidence sempre que não achar
            if button_location:
                print('img_lokan_relatorio localizado')
                return 'localizado'
        except:
            # py.press('esc',presses=2)
            print('tentando localizar - img_lokan_relatorio')
        time.sleep(3)  # Espera 1 segundo antes de tentar novamente
 


    
def verifica_vazio_notas_debitos():
    time.sleep(3)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\texto_vazio_notas_debitos.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                print('aui')
                return 'sim'
            
        except:
            try:
                button_location1920x1080 = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\texto_vazio_notas_debitos1920x1080.png", confidence=0.9) #alterar o confidence sempre que não achar
                
                if button_location1920x1080:
                    return 'sim'
            except:
                            # py.press('esc',presses=2)
                    print('tentando localizar - texto_vazio_notas_debitos')
                    return 'nao'
        time.sleep(2)  # Espera 3 segundos antes de tentar novamente




def verifica_texto_pagamentos():
    time.sleep(5)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\texto_disponivel_pagamentos.png", confidence=0.5) #alterar o confidence sempre que não achar
            if button_location:
                return 'localizado'
        except:

            try:
                button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\texto_disponivel_pagamentos1920x1080.png", confidence=0.9) #alterar o confidence sempre que não achar
                if button_location:
                    return 'localizado'
            except:
                # py.press('esc',presses=2)
                print('tentando localizar - texto_pagamentos')
                return 'seguir'

        time.sleep(2)  # Espera 3 segundos antes de tentar novamente
     


def verifica_texto_notas_canceladas():
    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\texto_disponivel_notas_canceladas.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                return 'sim'
        except:
            try:
                button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\texto_disponivel_notas_canceladas1920x1080.png", confidence=0.9) #alterar o confidence sempre que não achar
                if button_location:
                    return 'sim'
            except:
                 print('tentando localizar - texto_disponivel_notas_canceladas')
                 return 'nao'

            # py.press('esc',presses=2)
            # py.press('esc',presses=2)
             
        time.sleep(3)  #Espera 3 segundo antes de tentar novamente


 
       
def verifica_icone_relatorio():
    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\icone_relatorio.png", confidence=0.7) #alterar o confidence sempre que não achar
            if button_location:
                print('icone_relatorio localizado')
                return 'localizado'
        except:
            # py.press('esc',presses=2)
            # print('tentando localizar - icone_relatorio')
            return 'nao'
        time.sleep(3)  # Espera 3 segundo antes de tentar novamente





def entrada_usuario():
    time.sleep(3)  # Aguarda 1 segundo antes de iniciar

    tentativas = 0
    while tentativas != 15:
        try:
            # Localiza o ícone na tela com confiança de 50%
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\icon_usuario_loginV3.png", confidence=0.7) #aumentar confidence para ser mais preciso
            
            if button_location:
                # Pegando o centro da imagem encontrada
                # centro = pyautogui.center(button_location)
                
                # pyautogui.center(button_location)
                time.sleep(3)
                pyautogui.click(button_location,clicks=3)
                print('botaão entrada do usario localizado')

                # # Clicando 25 pixels à direita
                # pyautogui.click(centro.x + 25, centro.y,clicks=2)
                
                # print(f"Clicou na posição: {centro.x + 25}, {centro.y}")
                return 'localizado'
            # else:
            #     print("Tentando encontrar entrada de usuario - icon_usuario_login")
        except:
            tentativas += 1
            # py.hotkey('alt','tab',interval=1)
            # pyautogui.press('esc', presses=2)  # Pressiona 'esc' duas vezes se houver erro
            print('Botão não localizado - icon_usuario_login - tantativas: ',tentativas)
            time.sleep(1)  # Espera 5 segundo antes de tentar novamente
 
    return 'erro login'


def is_excel_open():
    """Verifica se o Microsoft Excel está aberto."""
    for process in psutil.process_iter(attrs=['name']):
        if process.info['name'] and "EXCEL.EXE" in process.info['name'].upper():
            return True
    return False

def verifica_icone_excel_superior():
    time.sleep(1)

    excel_pids = is_excel_open()
    
    if excel_pids:
        print(f"Excel está aberto. Aguardando 10 segundos para fechar...")
        time.sleep(10)  # Aguarda 10 segundos
        py.hotkey('alt','F4')
        print('Excel fechado')
        return 'localizado'
    else:
        print('Excel nao encontrado')
        return 'nao localizado'



def verifica_salvar_relatorio():
    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\icone_salvar_relatorio.png", confidence=0.2) #alterar o confidence sempre que não achar
            if button_location:
                print('texto salvar relatorio localizado')
                return 'localizado'
        except:
            # py.press('esc',presses=2)
            print('tentando localizar - texto salvar relatorio ')
        time.sleep(3)  # Espera 1 segundo antes de tentar novamente




def verifica_bem_vindo_login():
    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\login.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                print('texto bem vindo localizado ')
                return 'localizado'
        except:
            py.press('esc',presses=2)
            time.sleep(3)
            py.click(263,644)
            print('tentando localizar - texto bem vindo')
        time.sleep(10)  # Espera 1 segundo antes de tentar novamente
  

def btn_cancelar():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\btn_cancelar.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado - btn_cancelar')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - btn_cancelar')
        time.sleep(3)  # Espera 1 segundo antes de tentar novamente
  
def notas_fiscais():

    time.sleep(10)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\notas_fiscais.png", confidence=0.7) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location),clicks=2)
                print('botão  localizado - notas_fiscais')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - notas_fiscais')
        time.sleep(5)  # Espera 3 segundo antes de tentar novamente
 

def btn_relatorio_car():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\relatorio_car.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado - relatorio_car')
                return 'localizado'
        except:
            #py.press('esc',presses=2)
            print('botão nao localizado - relatorio_car')
        time.sleep(3)  # Espera 1 segundo antes de tentar novamente
     

def icon_contas_receber():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\icon_contas_receber.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                #py.press('esc',presses=2)
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado - icon_contas_receber')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - icon_contas_receber')
        time.sleep(5)  # Espera 1 segundo antes de tentar novamente
  
     


def icon_salvar():

    time.sleep(5)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\icon_salvar.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado - icon_salvar')
                break
        except:
            # py.press('esc',presses=2)
            print('botão nao localizado - icon_salvar')
        time.sleep(5)  # Espera 5 segundo antes de tentar novamente
  
        

def bnt_muda_empresa():

    time.sleep(1)
    tent = 0
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\bnt_muda_empresa.png", confidence=0.5) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado - bnt_muda_empresa')
                return 'sim'
                
        except:
            py.press('esc',presses=2)
            tent = tent + 1
            print(f'botão nao localizado - bnt_muda_empresa {tent}x')

            if tent >= 15:
                print('Erro ao tentar mudar de empresa')
                return 'nao'
        time.sleep(3)  # Espera 1 segundo antes de tentar novamente
  




def relatorio_faturamento_geral():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\relatorio_faturamento_geral.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location),clicks=2,interval=2)
                print('botão  localizado - relatorio_faturamento_geral')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - relatorio_faturamento_geral')
            return 'nao'
        time.sleep(5)  # Espera 3 segundo antes de tentar novamente
        


def btn_notas_canceladas():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\btn_notas_canceladas.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado - btn_notas_canceladas')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - btn_notas_canceladas')
        time.sleep(1)  # Espera 1 segundo antes de tentar novamente

def btn_relatorio_cap():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\relatorio_cap.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado - relatorio_cap')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - relatorio_cap')
        time.sleep(1)  # Espera 1 segundo antes de tentar novamente
        
def btn_icone_contas_a_pagar():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\icone_contas_a_pagar.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado - icone_contas_a_pagar')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - icone_contas_a_pagar')
        time.sleep(1)  # Espera 1 segundo antes de tentar novamente      

def btn_relatorios():

    time.sleep(1)
    tenta = 0
    while tenta != 10:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\btn_relatorios.png", confidence=0.7) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado - btn_relatorios')
                return 'localizado'
        except:
            # py.press('esc',presses=2)
            tenta = tenta + 1
            print(f'botão nao localizado - btn_relatorios | {tenta}x')
            
        time.sleep(3)  # Espera 1 segundo antes de tentar novamente

    return 'nao localizado'



def btn_copiar_produtos():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\copiar_produtos.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - btn_copiar_produtos')
        time.sleep(1)  # Espera 1 segundo antes de tentar novamente


def entrada_pesquisa_relatorio():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\entrada_pesquisa_relatorio.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado - entrada_pesquisa_relatorio')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - entrada_pesquisa_relatorio')
        time.sleep(1)  # Espera 1 segundo antes de tentar novamente





def btn_fechar_aba():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\btn_fechar_aba.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location),clicks=2,interval=1.5)
                print('botão  localizado - aba fechar')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - aba fechar')
        time.sleep(1)  # Espera 1 segundo antes de tentar novamente




def check_desmacar_todos():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\check_desmarcar_todos.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado check desmacar todos')
        time.sleep(1)  # Espera 1 segundo antes de tentar novamente
        
def check_num_nf():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\check_numnf.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location),clicks=2)
                print('botão  localizado')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - check numero')
        time.sleep(1)  # Espera 1 segundo antes de tentar novamente
        
        
def check_empresa():
    
    
    # Defina a região onde você quer procurar o texto à esquerda
    # x, y, largura, altura (ajuste de acordo com a sua tela)
    region_left = (259, 113, 280, 500)  # exemplo de coordenadas
    259,113
    553,108

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\check_empresa.png", confidence=0.9,region=region_left) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location),clicks=2)
                print('botão  localizado')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - check empresa')
        time.sleep(1)  # Espera 1 segundo antes de tentar novamente
        
        


def emitir():

    time.sleep(1)
    while True:
        try:
            
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\emitir.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado  - bnt emitir')
        time.sleep(1)  # Espera 1 segundo antes de tentar novamente
        
def btn_pesquisar():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\btn_pesquisar.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - pesquisar')
        time.sleep(3)  # Espera 1 segundo antes de tentar novamente
        

def bnt_selecionar_campos():

    time.sleep(1)
    while True:
        try:
            py.press('esc',presses=2)
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\bnt_selecionar_campos.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - selecionar campos')
        time.sleep(3)  # Espera 1 segundo antes de tentar novamente
        
def btn_abrir_origem():

    time.sleep(1)
    while True:
        try:
            py.press('esc',presses=2)
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\abrir_origem.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location))
                print('botão  localizado')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - btn_abrir_origem')
        time.sleep(3)  # Espera 1 segundo antes de tentar novamente
        
def nota_encontrada():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\nota_encontrada.png", confidence=0.9) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(pyautogui.center(button_location),clicks=2)
                print('botão  localizado')
                break
        except:
            py.press('esc',presses=2)
            print('botão nao localizado - nota_encontrada')
        time.sleep(3)  # Espera 1 segundo antes de tentar novamente


def verifica_texto_titulos_relatorio():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\texto_titulos_relatorio.png", confidence=0.5) #alterar o confidence sempre que não achar
            if button_location:
                centro = pyautogui.center(button_location)

                pyautogui.click(centro.x - 120,centro.y - 14) #provisorio

                pyautogui.click(centro.x - 120,centro.y) #aberto ( a vencer)

                pyautogui.click(centro.x - 120,centro.y + 14) #aberto (vencido)

                pyautogui.click(centro.x + 60,centro.y) #liquidado 

                print('texto localizado - verifica_texto_titulos_relatorio')
                break
        except:
            # py.press('esc',presses=2)
            print('botão nao localizado - verifica_texto_titulos_relatorio')
        time.sleep(3)  # Espera 1 segundo antes de tentar novamente




def verifica_texto_titulos_relatorio_aberto():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\texto_titulos_relatorio.png", confidence=0.5) #alterar o confidence sempre que não achar
            if button_location:
                centro = pyautogui.center(button_location)

                # pyautogui.click(centro.x + 60,centro.y) #liquidado 

                pyautogui.click(centro.x + 60,centro.y - 14) #liquidado a compensar

                # pyautogui.click(centro.x - 120,centro.y) #aberto ( a vencer)

                # pyautogui.click(centro.x - 120,centro.y - 14) #provisorio

                

                pyautogui.click(centro.x - 120,centro.y + 14) #aberto (vencido)

                


                print('texto localizado - verifica_texto_titulos_relatorio')
                break
        except:
            # py.press('esc',presses=2)
            print('botão nao localizado - verifica_texto_titulos_relatorio')
        time.sleep(3)  # Espera 1 segundo antes de tentar novamente



def verifica_texto_marcar_todos():

    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\texto_marcar_todos.png", confidence=0.7) #alterar o confidence sempre que não achar
            if button_location:
                centro = pyautogui.center(button_location)


                #687,442 resolução 1336x766 - lokan

                pyautogui.click(centro) #provisorio

                print('texto localizado - verifica_texto_marcar_todos')
                break
        except:
            # py.press('esc',presses=2)
            print('botão nao localizado - verifica_texto_marcar_todos')
        time.sleep(3)  # Espera 1 segundo antes de tentar novamente




def verifica_texto_notas_creditos():
    time.sleep(3)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\texto_disponivel_notas_creditos.png", confidence=0.7) #alterar o confidence sempre que não achar
            if button_location:
                return 'sim'
        except:
            try:
                button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\texto_disponivel_notas_creditos1920x1080.png", confidence=0.7) #alterar o confidence sempre que não achar
                if button_location:
                    return 'sim'
            except:
                print('tentando localizar - texto_disponivel_notas_creditos')
                return 'nao' 
             
        time.sleep(3)  #Espera 3 segundo antes de tentar novamente
       



def verifica_btn_iniciar():
    time.sleep(1)
    while True:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\btn_iniciar.png", confidence=0.5) #alterar o confidence sempre que não achar
            if button_location:
                # pyautogui.click(button_location)
                
                return 'sim'
        except:
            # py.press('esc',presses=2)
            # py.press('esc',presses=2)
            print('tentando localizar - btn_iniciar')
            
        time.sleep(3)  #Espera 3 segundo antes de tentar novamente



def verifica_btn_Ok():
    tent = 0
    while tent != 1:
        try:
            button_location = pyautogui.locateOnScreen(r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Imagens\bnt_ok.png", confidence=0.7) #alterar o confidence sempre que não achar
            if button_location:
                pyautogui.click(button_location)
                print('localizado - btn_OK_1366x768')
        except:
            # py.press('esc',presses=2)
            # py.press('esc',presses=2)
            tent += 1
            print(f'tentando localizar - btn_OK_1366x768 - {tent}x')
        time.sleep(3)  #Espera 3 segundo antes de tentar novamente





# print(verifica_texto_notas_creditos())

# check_desmacar_todos()
# check_empresa()
# check_num_nf()
# btn_ok()


