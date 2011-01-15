#!/usr/bin/python
# -*- coding: iso-8859-1 -*-

from xlrd import open_workbook, XL_CELL_BLANK, XL_CELL_EMPTY
from CCell import CCell
from time import gmtime,strftime

#Ano usado para acessar arquivos, planilhas, etc, que o inculam no nome
YEAR = strftime("%Y", gmtime())
#Nome da empresa usado de maneira parecida com o ano (YEAR)
COMPANY = ""

#Este array guarda os nomes das planilhas de CC
#MONTH_ALIAS_LIST = ["JAN-2011","FEV-2011","MAR-2011","ABR-2011","MAI-2011","JUN-2011",
#                    "JUL-2011","AGO-2011","SET-2011","OUT-2011","NOV-2011","DEZ-2011"]
MONTH_ALIAS_LIST = ["JAN-"+str(YEAR),"FEV-"+str(YEAR),"MAR-"+str(YEAR),"ABR-"+str(YEAR),
                    "MAI-"+str(YEAR),"JUN-"+str(YEAR),"JUL-"+str(YEAR),"AGO-"+str(YEAR),
                    "SET-"+str(YEAR),"OUT-"+str(YEAR),"NOV-"+str(YEAR),"DEZ-"+str(YEAR)]
#No Windows:
#FILE_CASH = ".\\livro_caixa\\CASHJEEP21122010.xls"
#FILE_CC = ".\\livro_caixa\\ccespecie24092010.xls"
#No Linux:
FILE_CASH = "./livro_caixa/CASH" + str(COMPANY) + str(YEAR) + ".xls"
DIR_FILES = "./livro_caixa/" 
FILE_MONEY = "./livro_caixa/ESPECIE" + str(YEAR) + ".xls"
FILE_CC = ""
#FILE_CC = ".\livro caixa\ESPECIE2011.xls"

#BOOK_CASH = open_workbook(filename=FILE_CASH,on_demand=True)
BOOK_CASH = open_workbook(filename=FILE_CASH)
#BOOK_CC = open_workbook(filename=FILE_CC,on_demand=True)
#BOOK_CC = open_workbook(filename=FILE_CC)
#Workbooks do solicitante ESPECIE
BOOK_MONEY = open_workbook(filename=FILE_MONEY)
#Lista dos workbooks dos solicitantes
BOOK_CC_LIST = []
#Lista dos solicitantes
REQUESTORS = []

#Planilha Conciliação, no arquivo CASH
CASH_SHEET_CONCILIATION = 'CONCILIAÇÃO'

#Planilha de Despesas detalhadas, no arquivo CASH
CASH_SHEET_DETAILED_EXPENSES = 'Despesas detalhadas'
CASH_SHEET_DETAILED_EXPENSES_COL_REQUESTOR = 3
CASH_SHEET_DETAILED_EXPENSES_COL_DUEDATE = 2
#Coluna "Código da Razão Geral"
CASH_SHEET_DETAILED_EXPENSES_COL_CODE = 0
#Colunas de chaves
CASH_SHEET_DETAILED_EXPENSES_COL_KEY1 = 10
CASH_SHEET_DETAILED_EXPENSES_COL_KEY2 = 11
CASH_SHEET_DETAILED_EXPENSES_COL_KEY3 = 12
CASH_SHEET_DETAILED_EXPENSES_COL_KEY4 = 13;

#Os arquivos de CC tem suas planilhas organizadas por meses do ano
CC_SHEET_COL_CODE = 1 #Plano de contas
CC_SHEET_COL_DATE = 2 #Data do vencimento 

def SetBookCASH(filepath):
    FILE_CASH = filepath
    
def SetBookCC(filepath):
    FILE_CC = filepath

def CompareMoney(a, b):
    #Se a ou b possui muitas casas decimais, uma maneira de trucar para apenas 2
    #é converter para string
    a = str(a)
    b = str(b)
    if a == b:
        return True
    else:
        return False

def GetRowIndexfromCol(sheet, col, value):
    #Dada uma coluna e a planilha, esta funcao busca o valor também dado nesta coluna da planilha
    for i in range(0, sheet.nrows):
        if sheet.cell(i, col).value == value:
            return i
    print "Não foi possível achar na planilha ", sheet.name, " na coluna ", col, " o valor ", value
    return -1

def GetColIndexfromRow(sheet, row, value):
    #Dada uma coluna e a planilha, esta funcao busca o valor também dado nesta coluna da planilha
    for i in range(0, sheet.ncols):
        if sheet.cell(row, i).value == value:
            return i
    print "Não foi possível achar na planilha ", sheet.name, " na linha ", row, " o valor ", value
    return -1

def GetLastColCell(sheet, col):
    last = sheet.nrows-1
    if last < 0:
        print "Erro: coluna ", col, " não tem valores"
    while last >= 0:
        cell = sheet.cell(last, col)
        if (cell.ctype == XL_CELL_EMPTY) or (cell.ctype == XL_CELL_BLANK):
            last = last-1
        else:
            ccell = CCell(cell)
            ccell.setPos(last, col)
            return ccell
    print "Erro: coluna ", col, " da planilha ", sheet, " n"
    exit()
