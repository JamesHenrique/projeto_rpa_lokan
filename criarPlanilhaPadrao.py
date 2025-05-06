import os
import glob
import time
import glob
from caminhos import CAMINHO_NOTAS_DEBITOS,CAMINHO_NOTAS_CREDITOS,CAMINHO_NOTAS_PAGAMENTOS,CAMINHO_NOTAS_CANCELADAS
from logger import infoLogs



def excluir_arquivos_xls():
    # Defina as pastas onde os arquivos .xls devem ser excluídos
    pastas = [
        CAMINHO_NOTAS_CREDITOS,
        CAMINHO_NOTAS_DEBITOS,
        CAMINHO_NOTAS_CANCELADAS,
        CAMINHO_NOTAS_PAGAMENTOS
    ]
    
    # Calcula o timestamp atual e a diferença de 3 dias
    agora = time.time()
    tres_dias_segundos = 3 * 24 * 60 * 60  # 3 dias em segundos
    
    for pasta in pastas:
        # Padrões separados para .xls e .xlsx
        arquivos_xls = glob.glob(os.path.join(pasta, '*.xls'))
        arquivos_xlsx = glob.glob(os.path.join(pasta, '*.xlsx'))
        arquivos = arquivos_xls + arquivos_xlsx
        
        for arquivo in arquivos:
            try:
                ultima_modificacao = os.path.getmtime(arquivo)
                if agora - ultima_modificacao > tres_dias_segundos:
                    os.remove(arquivo)
                    infoLogs().info(f"Arquivo {arquivo} excluído com sucesso.")
            except Exception as e:
                infoLogs().info(f"Erro ao excluir {arquivo}: {e}")