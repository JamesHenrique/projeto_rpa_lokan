import time as tm

import pyautogui

tm.sleep(3)

def clique_nota_debito():
   

    # Posição original do clique na resolução 1366x768
    base_x, base_y = 709,482

    # Resolução base da automação
    base_width, base_height = 1366, 768

    # Obtenha o tamanho da tela atual
    screen_width, screen_height = pyautogui.size()

    # Calcule a proporção de ajuste
    width_ratio = screen_width / base_width
    height_ratio = screen_height / base_height

    # Ajuste a posição conforme a resolução atual
    adjusted_x = int(base_x * width_ratio)
    adjusted_y = int(base_y * height_ratio)

    # Move e clica na posição ajustada
    pyautogui.moveTo(adjusted_x, adjusted_y)
    pyautogui.click()

    print(f"Clicou em: ({adjusted_x}, {adjusted_y})")


def verifica_tamanho_tela():

    # Obtém o tamanho da tela
    screen_width, screen_height = pyautogui.size()

    # Verifica se a resolução é 1920x1080
    if screen_width == 1920 and screen_height == 1080:
        print("A resolução da tela é 1920x1080.")
        return '1920x1080'
    elif screen_width == 1366 and screen_height == 768:
        print("A resolução da tela é 1366x768.")
        return '1366x768'
    
    else:
        return '1366x768'
        
        


