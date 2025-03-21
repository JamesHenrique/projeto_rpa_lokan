from datetime import datetime, timedelta

from datas import carregar_datas




# Exemplo de uso para acessar as datas
datas = carregar_datas()

periodo = ""



"""
Arquivo para deixar todos os caminhos relevantes para o funcionamento do bot
"""
mes = datetime.now().date().month
ano = datetime.now().date().year


match mes:
        case 1:
            mes = "Janeiro"
        case 2:
            mes = "Fevereiro"
        case 3:
            mes = "Março"
        case 4:
            mes = "Abril"
        case 5:
            mes = "Maio"
        case 6:
            mes = "Junho"
        case 7:
            mes = "Julho"
        case 8:
            mes = "Agosto"
        case 9:
            mes = "Setembro"
        case 10:
            mes = "Outubro"
        case 11:
            mes = "Novembro"
        case 12:
            mes = "Dezembro"




    #C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app
def planilhaPastaNotasCanceladas(empresa):
    # Exemplo de uso para acessar as datas
    datas = carregar_datas()

    periodo = ""
        
    if isinstance(datas, tuple):
        # Se forem duas datas, imprime as duas
        periodo = f'{datas[0]}_{datas[1]}'
        return rf"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilhas\Notas_debitos_canceladas\notas_canceladas_{periodo}_{empresa}.xlsx"
    elif datas:
        return rf"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilhas\Notas_debitos_canceladas\notas_canceladas_{datas}_{empresa}.xlsx"





def caminhoPlanilhaNotasDebito(empresa):
    # Exemplo de uso para acessar as datas
    datas = carregar_datas()

    periodo = ""
    if isinstance(datas, tuple):
        # Se forem duas datas, imprime as duas
        periodo = f'{datas[0]}_{datas[1]}'
        return rf"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilhas\Notas_debitos\notas_debitos_{periodo}_{empresa}.xlsx"   
    elif datas:
       return rf"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilhas\Notas_debitos\notas_debitos_{datas}_{empresa}.xlsx"
    


    


def caminhoPlanilhaPagamentosCAR(empresa):
    # Exemplo de uso para acessar as datas
    datas = carregar_datas()

    periodo = ""
    if isinstance(datas, tuple):
        # Se forem duas datas, imprime as duas
        periodo = f'{datas[0]}_{datas[1]}'
        return rf"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilhas\Pagamentos\pagamentos_{periodo}_{empresa}.xlsx"
    elif datas:
        return rf"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilhas\Pagamentos\pagamentos_{datas}_{empresa}.xlsx"


def caminho_planilha_notas_CAP():
    # Exemplo de uso para acessar as datas
    datas = carregar_datas()

    periodo = ""
    if isinstance(datas, tuple):
        # Se forem duas datas, imprime as duas
        periodo = f'{datas[0]}_{datas[1]}'
        return rf"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilhas\Notas_creditos\notas_creditos_{periodo}.xlsx"
                    
    elif datas:
        return rf"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilhas\Notas_creditos\notas_creditos_{datas}.xlsx"

def caminho_planilha_notas_CAP_aberto(empresa):
    # Exemplo de uso para acessar as datas
    datas = carregar_datas()

    periodo = ""
    if isinstance(datas, tuple):
        # Se forem duas datas, imprime as duas
        periodo = f'{datas[0]}_{datas[1]}'
        return rf"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilhas\Notas_creditos\notas_creditos_{periodo}_aberto_{empresa}.xlsx"
    elif datas:
        return rf"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilhas\Notas_creditos\notas_creditos_{datas}_aberto_{empresa}.xlsx"






def planilhDestino():
    return rf"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilha principal\planilha_Padrao_{mes}-{ano}.xlsx"

CAMINHO_NOTAS_CREDITOS = r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilhas\Notas_creditos"

CAMINHO_NOTAS_DEBITOS = r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilhas\Notas_debitos"

CAMINHO_NOTAS_CANCELADAS = r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilhas\Notas_debitos_canceladas"

CAMINHO_NOTAS_PAGAMENTOS = r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilhas\Pagamentos"

CAMINHO_DESTINO = r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Planilha principal"

CAMINHO_LOGS = r"C:\Users\Lokan Sisten\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Logs"

EMPRESAS = [1,2,3,4] #LOKAN - LOKANESTRUTURA - LOKAN FACIL - LOKANLONDRINA








