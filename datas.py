from datetime import datetime, timedelta
import os


ARQUIVO_DATA = r"C:\Users\lokan\OneDrive\Documentos\projeto_auto_Lokan -V2\app\Data\data_escolhida.txt"

# #data para gerar o relatório
def ontem():
    hoje = datetime.now().date()
    ontem = timedelta(days=1) #trocar data
    filtro = hoje - ontem
    formatoDia = filtro.strftime("%d-%m-%Y")
    return formatoDia




 #Função para salvar as datas escolhidas
def salvar_data(data):
    with open(ARQUIVO_DATA, 'w') as f:
        f.write(data)

# Função para carregar a data do arquivo
def carregar_data():
    if os.path.exists(ARQUIVO_DATA):
        with open(ARQUIVO_DATA, 'r') as f:
            return f.read().strip()
    return None
# Função para obter a data do relatório
def todasDatas():
    # Perguntar ao usuário se deseja usar o dia anterior ou inserir novas datas
    
    print('-------------------------- RPA LOKAN V2 --------------------------\n')
    print('Escolha abaixo se quer realizar a busca de ontem ou especificar um periodo!\n\n')
    escolha = input(f"Deseja usar a data de ontem ?\n Ontem {ontem()}\n Responda com: 'S' para sim e 'N' para não..: ").strip().lower()

    if escolha == 's':
        # Usar o dia anterior
        hoje = datetime.now().date()
        filtro = hoje - timedelta(days=1)
        formatoDia = filtro.strftime("%d-%m-%Y")
        salvar_data(formatoDia)  # Salvar data do dia anterior
        return formatoDia
    else:
        # Solicitar duas datas no formato dd/mm/yyyy
        while True:
            try:
                # Primeira data
                data_input1 = input("Digite a primeira data desejada (dd/mm/yyyy): ").strip()
                data1 = datetime.strptime(data_input1, "%d/%m/%Y").date()
                formatoDia1 = data1.strftime("%d-%m-%Y")
                
                # Segunda data
                data_input2 = input("Digite a segunda data desejada (dd/mm/yyyy): ").strip()
                data2 = datetime.strptime(data_input2, "%d/%m/%Y").date()
                formatoDia2 = data2.strftime("%d-%m-%Y")

                # Verifica se a segunda data é posterior à primeira
                if data2 >= data1:
                    break  # Sai do loop se ambas as datas forem válidas
                else:
                    print("A segunda data deve ser igual ou posterior à primeira.")
            except ValueError:
                print("Data inválida. Por favor, use o formato dd/mm/yyyy.")

        # Salvar ambas as datas no arquivo, separadas por uma vírgula
        salvar_data(f"{formatoDia1},{formatoDia2}")
        return formatoDia1, formatoDia2

# Função para carregar as datas do arquivo
def carregar_datas():
    if os.path.exists(ARQUIVO_DATA):
        with open(ARQUIVO_DATA, 'r') as f:
            data_salva = f.read().strip()  # Lê o conteúdo do arquivo e remove espaços em branco extras
            datas = data_salva.split(',')  # Divide as datas usando a vírgula como separador (caso tenha duas datas)
            
            if len(datas) == 2:
                # Se houver duas datas
                return datas[0].strip(), datas[1].strip()  # Retorna as duas datas como uma tupla
            elif len(datas) == 1:
                # Se houver apenas uma data
                return datas[0].strip()  # Retorna a única data
    return None  # Se o arquivo não existir ou estiver vazio, retorna None







