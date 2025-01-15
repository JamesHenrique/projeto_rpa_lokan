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
    tres_dias_segundos = 2 * 24 * 60 * 60  # 3 dias em segundos
    
    for pasta in pastas:
        # Cria o padrão para arquivos .xls
        padrao = os.path.join(pasta, '*.xls')
        # Localiza todos os arquivos .xls na pasta
        arquivos_xls = glob.glob(padrao)
        
        # Verifica a data de modificação e exclui os arquivos com mais de 3 dias
        for arquivo in arquivos_xls:
            try:
                ultima_modificacao = os.path.getmtime(arquivo)
                if agora - ultima_modificacao > tres_dias_segundos:
                    os.remove(arquivo)
                    infoLogs().info(f"Arquivo {arquivo} excluído com sucesso.")
            except Exception as e:
                infoLogs().info(f"Erro ao excluir {arquivo}: {e}")









excluir_arquivos_xls()
