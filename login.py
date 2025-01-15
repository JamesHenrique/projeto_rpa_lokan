import pyautogui as py
import time as tm
from logger import infoLogs
from acoesBotao import entrada_usuario
import subprocess
import psutil


"""

Passo 1 - ok

Faz login na plataforma

"""

# Função para abrir o programa
def abrir_sisloc():
    caminho_executavel = r"C:\Users\lokan\OneDrive\Documentos\Sisloc\Client\sisloc.exe"
    argumentos = "/Client"

    # Inicia o processo de forma assíncrona e retorna o objeto do processo
    processo = subprocess.Popen([caminho_executavel, argumentos])
    print("Sisloc iniciado.")
    
    return processo

# Função para fechar o programa
def fechar_sisloc(processo):
    if processo.poll() is None:  # Verifica se o processo ainda está rodando
        processo.terminate()     # Encerra o processo de forma graciosa
        print("Sisloc foi fechado.")
    else:
        print("Sisloc já havia sido encerrado.")
    


def login():
    
    infoLogs().info("ETAPA LOGIN INICIADA")
    
    #dados login
    login = "-"
    senha = "-"

    abrir_sisloc()

    ##Faz login no sistema
    entrada =  entrada_usuario()
    if 'erro login' in entrada:
        py.hotkey('alt','F4')
        abrir_sisloc()
    else:
        tm.sleep(1.5)
        py.click(681,404)
        tm.sleep(1.5)
        py.write(login)
        py.press('tab')
        tm.sleep(1.5)
        py.write(senha)
        tm.sleep(1.5)
        py.press('enter')


def fechar_sisloc():
    # Itera sobre todos os processos em execução
    for processo in psutil.process_iter(['pid', 'name']):
        # Verifica se o processo com o nome "sisloc.exe" está em execução
        if processo.info['name'] == "sisloc.exe":
            # Encerra o processo encontrado
            processo.terminate()
            print("Sisloc foi fechado.")
            return
    print("Sisloc já estava fechado ou não foi encontrado.")




