import logging
import os
from datetime import datetime
from caminhos import CAMINHO_LOGS
def infoLogs():
    # Criar o nome do arquivo de log com base na data atual
    data_atual = datetime.now().strftime("%d-%m-%Y")  # Exemplo: "2025-01-01"
    nome_arquivo_log = rf"{CAMINHO_LOGS}\log_{data_atual}.txt"

    # Configurando o logger
    logging.basicConfig(
        level=logging.DEBUG,  # Nível de log que você quer capturar
        format='%(asctime)s - %(levelname)s - %(message)s',  # Formato da mensagem de log
        datefmt='[%H:%M:%S] - %d-%m-%Y',  # Formato da data
        handlers=[
            logging.FileHandler(nome_arquivo_log, mode='a'),  # Salva os logs no arquivo txt, substituindo caso já exista
            logging.StreamHandler()  # Exibe os logs no console
        ]
    )
    logger = logging.getLogger(__name__)

    return logger
