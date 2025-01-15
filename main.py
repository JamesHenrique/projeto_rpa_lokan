from login import login
from gerarRelatorio import gerarRelatorio
from criarPlanilhaPadrao import excluir_arquivos_xls
from formatarPlanilhaExcel import formatacao_planilha_final


from logger import infoLogs
import time as tm
import pyautogui as py
import messagebox as ms
from datas import todasDatas
from login import fechar_sisloc

# from gerarRelatorioCAR import gerarRelatorioCarPG
# # from BuscarNotas import pesquisaNotas
# # from caminhos import caminhoPlanilhaNumeroDaNota, caminhoPlanilhaNotasDebito,planilhDestino




todasDatas()



login()

infoLogs().info("ETAPA LOGIN FINALIZADA")

tm.sleep(10)


gerarRelatorio()


py.press('esc',presses=10,interval=0.5)



fechar_sisloc()



formatacao_planilha_final()


excluir_arquivos_xls()


ms.showinfo("FINALIZADO RPA","Automação finalizada com sucesso")




