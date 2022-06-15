# -*- coding: utf-8 -*-
# ************************************************************
# Autor.......: Vladmir Cruz
# Data........: 06 de Junho de 2022
# Arquivo.....:
# Descricao...: Programa que le e escreve em arquivos Excel
# ************************************************************
# pip3 install requests openpyxl python-dotenv

# Imports
import openpyxl # biblioteca que faz a leitura e gravação em arquivos Excel
import os       # biblioteca que acessa funções do sistema operacional
import sys # biblioteca que acessa funções do sistema base, nesse caso nos retorna o sistema operacional
import shutil # biblioteca para manipular arquivos
import logging
import json
import requests
import os
from dotenv import load_dotenv


def createLogFile(log_file_name):
    if os.path.exists(log_file_name):
        os.remove(log_file_name)
    else:
        print("The file does not exist")            
    
    load_dotenv()
    logging.basicConfig(filename=log_file_name,
                        level=logging.DEBUG)        

def createResult(var_wks, var_loc, var_wkb, isMac):
    var_wkl = 1 # um contador de linhas para a planilha resultante
    # Preenchendo o cabeçalho da planilha resultante
    var_wks.cell(row=var_wkl, column=1).value = 'Nome do Arquivo'
    var_wks.cell(row=var_wkl, column=2).value = 'Mes'
    var_wks.cell(row=var_wkl, column=3).value = 'Ano'
    var_wks.cell(row=var_wkl, column=4).value = 'Nome do Professor'
    var_wks.cell(row=var_wkl, column=5).value = 'Contrato do Professor'
    var_wks.cell(row=var_wkl, column=6).value = 'Curso'
    var_wks.cell(row=var_wkl, column=7).value = 'Turma'
    var_wks.cell(row=var_wkl, column=8).value = 'Fase'
    var_wks.cell(row=var_wkl, column=9).value = 'Atividade'
    var_wks.cell(row=var_wkl, column=10).value = 'Data'
    var_wks.cell(row=var_wkl, column=11).value = 'Quantidade de Horas'
    var_wks.cell(row=var_wkl, column=12).value = 'Observacao'
    logging.debug('BEFORE SAVE ' + var_loc)

    if(isMac):
        var_wkb.save(var_loc + '/resultado/' + 'resultado.xlsx') # salva o workbook
    else:
        var_wkb.save(var_loc + '\\resultado\\' + 'resultado.xlsx') # salva o workbook

def sendPowerBI(array_send,wb):
    logging.debug("Array to Send")
    logging.debug(array_send)
    res = json.dumps(array_send)
    logging.debug(res)
    r = requests.post(os.environ['URL']+ os.environ['KEY'], data=res)

    logging.debug("Sending to Power BI")
    logging.debug("File  = " + str(wb))
    logging.debug("Code  = " + str(r.status_code))
    logging.debug("Info  = " + str(r.text))
    print("Sending to Power BI - Code " + str(r.status_code))
    print("Sending to Power BI - Code " + str(r.text))


# Variáveis
var_loc = os.getcwd() # caminho do diretório pai do script
var_loc_read = os.getcwd()+"/original"  # caminho do diretório de leitura
var_loc_done = os.getcwd()+"/lidos/"  # caminho do diretório de leitura
var_dir = os.listdir(var_loc_read) # acessando o diretório que contêm os arquivos
var_os = sys.platform # sistema operacional que o programa está rodando
var_flt = [] # lista de arquivos que serão lidos para o conteúdo 
var_mon = '' # mês que está sendo reportado
var_yea = '' # ano que está sendo reportado
var_prf = '' # nome do professor que está reportando
var_tct = '' # tipo de contrato do professor
var_crs = '' # curso que o professor está reportando
var_trm = '' # turma que o professor está reportando
var_fas = '' # fase do curso que o professor está reportando
var_atv = '' # tipo de atividade que o professor está reportando
var_dat = '' # data que ocorreu o evento sendo reportado
var_qhr = '' # quantidade de horas que o professor está reportando
var_obs = '' # observação que o professor possa fazer
var_naq = '' # variável para armazenar o nome do arquivo atual
var_col = 0 # um contador genérico para contar as colunas
var_lin = 0 # um contador genérico para contar as linhas
var_ctr = 0 # um contador genérico, afinal, todo programa precisa de um
var_wkb = openpyxl.Workbook()
var_wks = var_wkb.active
var_wkl = 0 # um contador de linhas para o Workbook

createLogFile('reading.log')


# verifica se os diretorios resultantes existem, se não, os cria
if(var_os == 'linux' or var_os == 'linux2' or var_os == 'darwin'): # verifica se o sistema operacional é Linux ou MacOS
    logging.debug('IS MAC')
    if(os.path.isdir(var_loc + '/lidos/') == False):
        os.mkdir(var_loc + '/lidos/') # diretorio que recebera os arquivos que foram lidos
    if(os.path.isdir(var_loc + '/resultado/') == False): # se não existe o diretório resultado
        os.mkdir(var_loc + '/resultado/') # cria diretorio que contem o arquivo resultante
  
    if(os.path.isfile(var_loc + '/resultado/' + 'resultado.xlsx') == False):
        createResult(var_wks, var_loc, var_wkb, True)
    else:       
        logging.debug(os.path.isfile(var_loc + '/resultado/' + 'resultado.xlsx'))
        logging.debug('AQUI ' + var_loc)
        var_wkb = openpyxl.load_workbook(var_loc + '/resultado/' + 'resultado.xlsx')
        var_wks = var_wkb.active # criando a sheet que receberá os dados
        var_wkl = var_wks.max_row

elif var_os == 'win32': # verifica se o sistema operacional é Windows
    if(os.path.isdir(var_loc + '\\lidos\\') == False):
        os.mkdir(var_loc + '\\lidos\\') # cria o diretorio que recebera os arquivos que foram lidos
    if(os.path.isdir(var_loc + '\\resultado\\') == False):
        os.mkdir(var_loc + '\\resultado\\') # cria o diretorio que contem o arquivo resultante
    
    if(os.path.isfile(var_loc + '\\resultado\\' + 'resultado.xlsx') == False):
        createResult(var_wks, var_loc, var_wkb, False)
    else:        
        var_wkb = openpyxl.load_workbook(var_loc + '\\resultado\\' + 'resultado.xlsx')
        var_wks = var_wkb.active # criando a sheet que receberá os dados
        var_wkl = var_wks.max_row


for files in var_dir: # lendo todos os arquivos que estão armazenados no diretório
    if(files.endswith('.xls') or files.endswith('.xlsx')):
        if(var_os == 'linux' or var_os == 'linux2' or var_os == 'darwin'): # verifica se o sistema operacional é Linux ou MacOS
            var_flt.append(var_loc_read + '/' + files) # grava o nome do arquivo no array de nomes
        elif var_os == 'win32': # verifica se o sistema operacional é Windows
            var_flt.append(var_loc_read + '\\' + files) # grava o nome do arquivo no array de nomes

logging.debug("ARQUIVOS Excel a ler")
logging.debug(var_flt)
# Iterando sobre a lista de nomes de arquivos e seu conteúdo
for wb in var_flt:

    array_send = []
    var_naq = var_flt[var_ctr]
    workbook = openpyxl.load_workbook(wb) # chamada que abre o arquivo desejado
    worksheet = workbook.worksheets[1] # definindo qual o caderno do Excel que será utilizado (nesse caso é o segundo, contagem começa em 0)

    var_mon = worksheet.cell(row=3, column=2).value # lendo o mês do relatório
    var_yea = worksheet.cell(row=3, column=6).value # lendo o ano do relatório
    var_prf = worksheet.cell(row=4, column=3).value # lendo o nome do professor
    var_tct = worksheet.cell(row=5, column=3).value # lendo o tipo de contrato do professor

    var_lin = 8 # contador começa em 8 por ser a primeira linha que contém informações sobre as atividades
    var_col = 1 # contador começa em 1 por ser a primeira coluna que contém informações sobre as atividades
    ind_array = 0
    while var_lin < 70: # iterando sobre a tabela com os dados das atividades reportadas        
        var_check_end = worksheet.cell(row=var_lin, column=6).value # lendo o curso que o professor indicou        
        if (isinstance(var_check_end, str) and "SUM" in var_check_end): # Checa se o cursor chegou no final da planilha
            var_lin = 100
        else:
            var_crs = worksheet.cell(row=var_lin, column=var_col).value # lendo o curso que o professor indicou
            if(var_crs is not None): # verifica se há um curso sendo reportado, em caso de em branco, pula
                var_trm = worksheet.cell(row=var_lin, column=var_col + 1).value # lendo a turma
                var_fas = worksheet.cell(row=var_lin, column=var_col + 2).value # lendo a fase
                var_atv = worksheet.cell(row=var_lin, column=var_col + 3).value # lendo a atividade
                var_dat = worksheet.cell(row=var_lin, column=var_col + 4).value # lendo a data da ocorrência
                var_qhr = worksheet.cell(row=var_lin, column=var_col + 5).value # lendo a quantidade de horas indicada
                var_obs = worksheet.cell(row=var_lin, column=var_col + 6).value # lendo a observação que foi apontada
                var_col = 1 # resetando o contador de colunas
                var_lin = var_lin + 1 # pulando para a próxima linha

                line_send = {}
                line_send['nome_arquivo'] = str(var_naq)
                line_send['Mes'] =str(var_mon)
                line_send['Ano'] = str(var_yea)
                line_send['Professor'] = str(var_prf)
                line_send['Contrato'] = str(var_tct)
                line_send['Curso'] =str(var_crs)
                line_send['Turma'] =str(var_trm)
                line_send['Fase'] = str(var_fas)
                line_send['Atividade'] = str(var_atv)
                line_send['Data'] = str(var_dat)
                line_send['Quantidade'] =str(var_qhr)
                line_send['Observacao'] = str(var_obs)

                array_send.append(line_send)

                var_wkl = var_wkl + 1

                var_wks.cell(row=var_wkl, column=1).value = str(var_naq)
                var_wks.cell(row=var_wkl, column=2).value = str(var_mon)
                var_wks.cell(row=var_wkl, column=3).value = str(var_yea)
                var_wks.cell(row=var_wkl, column=4).value = str(var_prf)
                var_wks.cell(row=var_wkl, column=5).value = str(var_tct)
                var_wks.cell(row=var_wkl, column=6).value = str(var_crs)
                var_wks.cell(row=var_wkl, column=7).value = str(var_trm)
                var_wks.cell(row=var_wkl, column=8).value = str(var_fas)
                var_wks.cell(row=var_wkl, column=9).value = str(var_atv)
                var_wks.cell(row=var_wkl, column=10).value = str(var_dat)
                var_wks.cell(row=var_wkl, column=11).value = str(var_qhr)
                var_wks.cell(row=var_wkl, column=12).value = str(var_obs)

                ind_array += 1
            else:
                var_lin = var_lin + 1 # pulando para a próxima linha
    
    sendPowerBI(array_send, wb) #Envia para o PowerBI
    array_send.clear() # Limpa o array
    shutil.move(wb, var_loc_done) #Move para a pasta de lidos
    var_ctr = var_ctr + 1 # pula para a próxima planiha quando termina a iteração da atual

var_wkb.save(var_loc + '/resultado/' + 'resultado.xlsx')



