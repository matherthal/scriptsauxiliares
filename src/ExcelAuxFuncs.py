#!/usr/bin/python
# -*- coding: iso-8859-1 -*-

from xlrd import open_workbook, XL_CELL_BLANK, XL_CELL_EMPTY
from CCell import ComplCell

#Este array guarda os nomes das planilhas de CC
MONTH_ALIAS_LIST = ["JANE2010","FEVER2010","MARÇ2010","ABRIL2010","MAIO2010",
                  "JUNHO2010","JULHO2010","AGOSTO2010"]
FILE_CASH = ".\\livro caixa\\CASHJEEP21122010.xls"
FILE_CC = ".\\livro caixa\\ccespecie24092010.xls"
BOOK_CASH = open_workbook(filename=FILE_CASH,on_demand=True)
BOOK_CC = open_workbook(filename=FILE_CC,on_demand=True)

CASH_SHEET = 'CONCILIAÇÃO'

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
            ccell = ComplCell(cell)
            ccell.setPos(last, col)
            return ccell
    print "Erro: coluna ", col, " da planilha ", sheet, " n"
    exit()
