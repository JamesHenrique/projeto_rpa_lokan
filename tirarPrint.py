# import pyautogui
# import time as tm
# import cv2
# import pytesseract
# import re


# tm.sleep(2)

# #Defina as coordenadas da área que você quer capturar (x, y, largura, altura)

# def printVerificaLogin():

#     x = 107
#     y = 131
#     largura = 160
#     altura = 15

#     # Capturar a área específica
#     screenshot = pyautogui.screenshot(region=(x, y, largura, altura))

#     # Salvar a captura de tela
#     screenshot.save('C:\\Users\\TI\\Documents\\James\\projeto_auto_Lokan\\app\\Prints\\login.png ')



# def verificaTextoLogin():
#     # Carregar a imagem'
#     imagem = cv2.imread(r"Prints/login.png")

#     #configuração Tesseract
#     caminho = r"C:\Users\TI\AppData\Local\Programs\Tesseract-OCR"
#     pytesseract.pytesseract.tesseract_cmd = caminho + r"\tesseract.exe"

#     # Converter para escala de cinza
#     imagem_cinza = cv2.cvtColor(imagem, cv2.COLOR_BGR2GRAY)

#     # Aplicar filtro de desfoque
#     imagem_desfocada = cv2.GaussianBlur(imagem_cinza, (1, 1), 0)

#     # Aplicar OCR na imagem
#     texto = pytesseract.image_to_string(imagem_desfocada)
    
#     if 'Bem-vindo ao SISLOC®.' in texto:
#         texto = "login efetuado com sucesso"
#     else:
#         texto = "erro ao efetuar login"
    
#     return texto

 
# # def tirarPrintNotasCanceladas():
    
# #     x = 146
# #     y = 248
# #     largura = 190
# #     altura = 15

# #     # Capturar a área específica
# #     screenshot = pyautogui.screenshot(region=(x, y, largura, altura))

# #     # Salvar a captura de tela
# #     screenshot.save('Prints/printNotasCanceladas.png')
    



 



