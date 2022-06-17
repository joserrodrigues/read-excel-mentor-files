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
from openpyxl.worksheet.table import Table, TableStyleInfo
import os       # biblioteca que acessa funções do sistema operacional
import sys # biblioteca que acessa funções do sistema base, nesse caso nos retorna o sistema operacional
import shutil # biblioteca para manipular arquivos
import logging
import json
import requests
import os
from dotenv import load_dotenv
import datetime


def createAppFiles(log_file_name, isMac):
    if os.path.exists(log_file_name):
        os.remove(log_file_name)
    else:
        print("The file does not exist")            
    
    load_dotenv()
    logging.basicConfig(filename=log_file_name, level=logging.DEBUG)        

    readFolder = var_loc + '\\lidos\\'
    resultFolder = var_loc + '\\resultado\\'
    if(isMac):
        readFolder = var_loc + '/lidos/'
        resultFolder = var_loc + '/resultado/'

    if(os.path.isdir(readFolder) == False):
        os.mkdir(readFolder) # diretorio que recebera os arquivos que foram lidos
    if(os.path.isdir(resultFolder) == False):  # se não existe o diretório resultado
        os.mkdir(resultFolder) # cria diretorio que contem o arquivo resultante

    if os.path.exists(getResultFile(isMac)):
        os.remove(getResultFile(isMac))

def getResultFile(isMac):    
    if(isMac):
        return var_loc + '/resultado/' + 'resultado.xlsx'        
    else:
        return var_loc + '\\resultado\\' + 'resultado.xlsx'


def createResult(var_loc, isMac):
    print("Criando Restul")
    wb = openpyxl.Workbook()    
    mainWks = wb.active
    
    # Preenchendo o cabeçalho da planilha resultante
    mainWks.cell(row=1, column=1).value = 'Nome do Arquivo'
    mainWks.cell(row=1, column=2).value = 'Mes'
    mainWks.cell(row=1, column=3).value = 'Ano'
    mainWks.cell(row=1, column=4).value = 'Nome do Professor'
    mainWks.cell(row=1, column=5).value = 'Contrato do Professor'
    mainWks.cell(row=1, column=6).value = 'Curso'
    mainWks.cell(row=1, column=7).value = 'Turma'
    mainWks.cell(row=1, column=8).value = 'Fase'
    mainWks.cell(row=1, column=9).value = 'Atividade'
    mainWks.cell(row=1, column=10).value = 'Data'
    mainWks.cell(row=1, column=11).value = 'Quantidade de Horas'
    mainWks.cell(row=1, column=12).value = 'Observacao'
    logging.debug('BEFORE SAVE ' + var_loc)

    for courses in array_courses:
        wks = wb.create_sheet(title=courses)
        wks.cell(row=1, column=1).value = 'EMPRESA'
        wks.cell(row=1, column=2).value = 'CAMPUS'
        wks.cell(row=1, column=3).value = 'TIPO'
        wks.cell(row=1, column=4).value = 'UNIDADE / DEPARTAMENTO'
        wks.cell(row=1, column=5).value = 'SEGMENTO'
        wks.cell(row=1, column=6).value = 'DETALHE (CURSO)'
        wks.cell(row=1, column=7).value = 'Tipo Contrato'
        wks.cell(row=1, column=8).value = 'Nº Matrícula (CLT) | Contato (PJ/RPA)'
        wks.cell(row=1, column=9).value = 'Responsável - Nome Completo'
        wks.cell(row=1, column=10).value = 'Modalidade'
        wks.cell(row=1, column=11).value = 'Turma'
        wks.cell(row=1, column=12).value = 'Disciplina / Fase'
        wks.cell(row=1, column=13).value = 'Num. Cap.'
        wks.cell(row=1, column=14).value = 'Capítulo'
        wks.cell(row=1, column=15).value = 'DATA (Entrega / Correção)'
        wks.cell(row=1, column=16).value = 'CONTROLE INTERNO - LAUDAS / HORAS (não usar)'
        wks.cell(row=1, column=17).value = 'Núm. págs / horas'
        wks.cell(row=1, column=18).value = 'Data de cadastro (interno)'        

    wb.save(getResultFile(isMac))  # salva o workbook


def mountInfo(wb, array_send, line_controller, fileName, month, year, prof, contract, course, class_txt, phase, activity, 
                final_data, final_machine_data, hours, obs):
    line_send = {}
    line_send['nome_arquivo'] = str(fileName)
    line_send['Mes'] =str(month)
    line_send['Ano'] = str(year)
    line_send['Professor'] = str(prof)
    line_send['Contrato'] = str(contract)
    line_send['Curso'] =str(course)
    line_send['Turma'] = str(class_txt)
    line_send['Fase'] = str(phase)
    line_send['Atividade'] = str(activity)
    line_send['Data'] = str(final_machine_data)
    line_send['Quantidade'] = str(hours)
    line_send['Observacao'] = str(obs)

    array_send.append(line_send)

    wkLine = line_controller['all']

    work_sheet = wb['Sheet']

    work_sheet.cell(row=wkLine, column=1).value = "VSTP"
    work_sheet.cell(row=wkLine, column=2).value = "FIAP ON"
    work_sheet.cell(row=wkLine, column=3).value = "EDUCACIONAL / PROFESSORES"
    work_sheet.cell(row=wkLine, column=4).value = "MBA ON"
    work_sheet.cell(row=wkLine, column=5).value = "Especialização"
    work_sheet.cell(row=wkLine, column=6).value = str(course).upper()
    work_sheet.cell(row=wkLine, column=7).value = "**Copiar"
    work_sheet.cell(row=wkLine, column=8).value = "**Copiar"
    work_sheet.cell(row=wkLine, column=9).value = str(prof).upper()
    work_sheet.cell(row=wkLine, column=10).value = str(activity)
    work_sheet.cell(row=wkLine, column=11).value = str(class_txt).replace(" ", "")
    work_sheet.cell(row=wkLine, column=12).value = str(phase)
    work_sheet.cell(row=wkLine, column=13).value = ""
    work_sheet.cell(row=wkLine, column=14).value = ""
    work_sheet.cell(row=wkLine, column=15).value = str(final_data)
    work_sheet.cell(row=wkLine, column=16).value = "**Copiar"
    work_sheet.cell(row=wkLine, column=17).value = str(hours).replace(".", ",")
    work_sheet.cell(row=wkLine, column=18).value = ""

    wkLine = line_controller[course]
    work_sheet = wb[course]
    work_sheet.cell(row=wkLine, column=1).value = "VSTP"
    work_sheet.cell(row=wkLine, column=2).value = "FIAP ON"
    work_sheet.cell(row=wkLine, column=3).value = "EDUCACIONAL / PROFESSORES"
    work_sheet.cell(row=wkLine, column=4).value = "MBA"
    work_sheet.cell(row=wkLine, column=5).value = "Especialização"
    work_sheet.cell(row=wkLine, column=6).value = str(course).upper()
    work_sheet.cell(row=wkLine, column=7).value = "**Copiar"
    work_sheet.cell(row=wkLine, column=8).value = "**Copiar"
    work_sheet.cell(row=wkLine, column=9).value = str(prof).upper()
    work_sheet.cell(row=wkLine, column=10).value = str(activity)
    work_sheet.cell(row=wkLine, column=11).value = str(class_txt).replace(" ", "")
    work_sheet.cell(row=wkLine, column=12).value = str(phase)
    work_sheet.cell(row=wkLine, column=13).value = ""
    work_sheet.cell(row=wkLine, column=14).value = ""
    work_sheet.cell(row=wkLine, column=15).value = str(final_data)
    work_sheet.cell(row=wkLine, column=16).value = "**Copiar"
    work_sheet.cell(row=wkLine, column=17).value = str(hours).replace(".", ",")
    work_sheet.cell(row=wkLine, column=18).value = ""

    line_controller['all'] += 1
    line_controller[course] += 1

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
var_wkl = 0 # um contador de linhas para o Workbook
line_controller = {}
array_courses = ['GTIO', 'AOJO', 'ASOO', 'ABDO', 'DTSO', 'BDTO', 'DGO', 'NGO', 'BIO', 'SGO', 'SCJO'] 

isMac = False
# verifica se o sistema operacional é Linux ou MacOS
if(var_os == 'linux' or var_os == 'linux2' or var_os == 'darwin'):
    print('Is Mac')    
    isMac = True

#cria os arquivos e pastas utilizadas no app
createAppFiles('reading.log', isMac)
logging.debug("Iniciando a Importacao")

# verifica se os diretorios resultantes existem, se não, os cria
if(os.path.isfile(getResultFile(isMac)) == False):
    #Cria a planilha de resultado    
    createResult(var_loc, isMac)
    
#abre a planilha de resultado
var_wkb = openpyxl.load_workbook(getResultFile(isMac))
var_wks = var_wkb.active  # criando a sheet que receberá os dados


for files in var_dir: # lendo todos os arquivos que estão armazenados no diretório
    if(files.endswith('.xls') or files.endswith('.xlsx')):
        if(var_os == 'linux' or var_os == 'linux2' or var_os == 'darwin'): # verifica se o sistema operacional é Linux ou MacOS
            var_flt.append(var_loc_read + '/' + files) # grava o nome do arquivo no array de nomes
        elif var_os == 'win32': # verifica se o sistema operacional é Windows
            var_flt.append(var_loc_read + '\\' + files) # grava o nome do arquivo no array de nomes

#mount lines Controller
line_controller['all'] = 1
for course in array_courses:
    line_controller[course] = 2

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

                #Trabalhando com a data 
                final_var_data = ""
                final_machine_data = ""
                if (isinstance(var_dat, datetime.datetime)):
                    final_var_data = var_dat.strftime("%d/%m/%Y")
                    final_machine_data = var_dat.strftime("%Y-%m-%d")
                elif (isinstance(var_dat, str)):
                    final_var_data = final_machine_data = var_dat[0: 10]
                else:
                    final_var_data = final_machine_data = var_dat

                mountInfo(var_wkb, array_send, line_controller, var_naq, var_mon, var_yea, var_prf, var_tct, var_crs,
                          var_trm, var_fas, var_atv, final_var_data, final_machine_data, var_qhr, var_obs)
            else:
                var_lin = var_lin + 1 # pulando para a próxima linha
    
    sendPowerBI(array_send, wb) #Envia para o PowerBI
    array_send.clear() # Limpa o array
    shutil.move(wb, var_loc_done) #Move para a pasta de lidos
    var_ctr = var_ctr + 1 # pula para a próxima planiha quando termina a iteração da atual

var_wkb.save(var_loc + '/resultado/' + 'resultado.xlsx')



