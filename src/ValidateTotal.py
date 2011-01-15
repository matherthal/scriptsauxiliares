#!/usr/bin/python
# -*- coding: iso-8859-1 -*-

from xlrd import XL_CELL_NUMBER,XL_CELL_DATE
from CCell import CCell
from ExcelAuxFuncs import *
import CCell
from time import gmtime,strftime,strptime

def StartValidade():
    YEAR = strftime("%Y", gmtime())
    COMPANY = "JEEPTOUR"
    #BOOK_CC_LIST = 

def SetRequestors():
    '''
    Coloca o nomes dos solicitantes em REQUESTORS.
    Inicia os workbooks de cada solicitante, colocando em BOOK_CC_LIST
    '''
    
    try:
        #Pegando a planilha de despesas detalhadas
        sheetCASH = BOOK_CASH.sheet_by_name(CASH_SHEET_DETAILED_EXPENSES)
    except:
        print "Erro: Não foi encontrada planilha ", CASH_SHEET_DETAILED_EXPENSES
        return 1
    
    #Pegando o nome dos solicitantes (requestor) em Despesa de detalhada
    COL_REQ = CASH_SHEET_DETAILED_EXPENSES_COL_REQUESTOR
    
    #Preencher as listas REQUESTORS
    for row in range(sheetCASH.nrows):
        #Se o solicitante não estiver na lista, inseri-lo
        cellreq = sheetCASH.cell(row, COL_REQ)
        if not(cellreq.value in REQUESTORS):
            #Lista de Solicitantes
            REQUESTORS.append(cellreq.value)
            #Dado o solicitante, pegar o workbook
            try:            
                filereq = DIR_FILES + cellreq.value + str(YEAR) + ".xls"
                workbookCC = open_workbook(filename=filereq)
                #Lista dos workbooks dos solicitantes
                BOOK_CC_LIST.append(workbookCC)
            except:
                print "Erro: Não foi encontrado o arquivo ", filereq
                return 1
    return 0

def ValidateTotal(ROWS_CASH, COL_CC):
    '''
    Valida o TOTAL da Receita, Despesa ou Saldo
    
    '''
    
    print 'Comparando arquivos: ', FILE_CASH, ' e ', FILE_CC
    try:
        sheetCASH = BOOK_CASH.sheet_by_name(CASH_SHEET_CONCILIATION)
    except:
        print "Erro: Não foi encontrada planilha CONCILIAÇÃO"
        return 1
    
    #Imaginando que cada arquivo só contém os meses de 1 ano.
    #Esta repetição percorre cada sheet de CC e cada coluna de mês de CASH
    for i in range(0,12): 
        if (len(MONTH_ALIAS_LIST)-1) < i:
            return 1          
        #Pegando a coluna do mês de CASH. A partir da coluna C.
        colCASH = i + 1; #Para descontar a primeira coluna
        val = 0
        for j in range(0,len(ROWS_CASH)):
            cellCASH = sheetCASH.cell(ROWS_CASH[j], colCASH)
            #print 'Comparando planilha ', MONTH_ALIAS_LIST[i], ' com coluna ', colCASH
            #print 'Célula (', ROW_CASH + 1, ',', colCASH + 1, ') = ', cellCASH.value
            if cellCASH.ctype==XL_CELL_NUMBER:
                val += cellCASH.value
            #else:
            #    print "Erro na planilha de CASH: A célula da linha ", ROWS_CASH[j], ' e coluna ', colCASH, ' não é um número.'
            #DEBUG print ROWS_CASH[j], ',', colCASH, ': ', str(cellCASH.value), ' tipo: ', cellCASH.ctype 
        
        #Pegando a sheet do mês de CC especie
        #A coluna é constante (G) e a linha é a última
        sheetCC = BOOK_MONEY.sheet_by_name(MONTH_ALIAS_LIST[i])
        #GetLastColCell retorna uma CCell
        ccellCC = GetLastColCell(sheetCC, COL_CC)
        rowCC = ccellCC.getRow(); colCC = ccellCC.getCol();
        #print 'Célula (', rowCC + 1, ',', colCC + 1, ') = ', ccellCC.getValue()
        if not(ccellCC.getType()==XL_CELL_NUMBER):
            print "Erro na planilha ", MONTH_ALIAS_LIST[i], " de CC: A célula da linha ", rowCC, ' e coluna ', colCC, ' não é um número.'
        
        if CompareMoney(ccellCC.getValue(), val):
            print "Total de ", MONTH_ALIAS_LIST[i], " ----------------------- OK"
        else:
            print "Total de ", MONTH_ALIAS_LIST[i], " ----------------------- ERRO"
    return 0

def ValidateTotalIncome():
    '''
    Valida o TOTAL da Receita
    
    '''
    print '1. VALIDAÇÃO DO TOTAL DA RECEITA'
    #Coluna da Receita na planilha de CC. Coluna J. É constantes na busca.
    COL_CC = 6
    #COL_CC = 9
    #Linhas importantes para cálculo da receita na planilha de CASH. É constantes na busca. 
    ROWS_CASH = [4,17,18,19,20]
    ValidateTotal(ROWS_CASH, COL_CC)
    print 

def ValidateTotalExpense():
    '''
    Valida o TOTAL da Despesa
    
    '''
    
    print '2. VALIDAÇÃO DO TOTAL DA DESPESA'
    #Coluna da Despesa na planilha de CC. Coluna I. É constantes na busca.
    COL_CC = 5
    #COL_CC = 8
    #Linhas importantes para cálculo da despesa na planilha de CASH. É constantes na busca. 
    ROWS_CASH = [6,24,25,26,27,28,29]
    ValidateTotal(ROWS_CASH, COL_CC)
    print 

def ValidateTotalBalance():
    '''
    Valida o TOTAL do Saldo
    
    '''
    
    print '3. VALIDAÇÃO DO SALDO'
    #Coluna do Saldo na planilha de CC. Coluna K. É constantes na busca.
    COL_CC = 7
    #COL_CC = 10
    #Linha do Saldo 
    ROWS_CASH = [31]
    ValidateTotal(ROWS_CASH, COL_CC)
    print

def FindErrorConstructionKeyExistence():
    '''
    Busca erro de não existência de chave para um código. Na planilha Despesas Detalhadas, arquivo CASH 
    Retorna uma tupla correspondente à posição da célula de erro. Senão retorna None.
    
    '''
    
    print 'BUSCANDO ERRO DE CHAVE NA PLANILHA "', CASH_SHEET_DETAILED_EXPENSES, '" em CASH.'
    try:
        sheetCASH = BOOK_CASH.sheet_by_name(CASH_SHEET_DETAILED_EXPENSES)
    except:
        print "Erro: Não foi encontrada planilha ", CASH_SHEET_DETAILED_EXPENSES
        return 1
    
    #Coluna "Código da Razão Geral"
    COL_CODE = CASH_SHEET_DETAILED_EXPENSES_COL_CODE
    #Colunas de chaves
    COL_KEY1 = CASH_SHEET_DETAILED_EXPENSES_COL_KEY1
    COL_KEY2 = CASH_SHEET_DETAILED_EXPENSES_COL_KEY2
    COL_KEY3 = CASH_SHEET_DETAILED_EXPENSES_COL_KEY3
    COL_KEY4 = CASH_SHEET_DETAILED_EXPENSES_COL_KEY4
    
    for row in range(sheetCASH.nrows):
        cellcode = sheetCASH.cell(row, COL_CODE)
        #Se existir o código da coluna A, e não existir uma das chaves das últimas colunas: Erro
        if cellcode.ctype == XL_CELL_NUMBER:
            #Pesquisando nas 4 chaves da linha. Se alguma estiver em branco ou vazia, returna sua posição
            cellkey1 = sheetCASH.cell(row, COL_KEY1)
            cellkey2 = sheetCASH.cell(row, COL_KEY2)
            cellkey3 = sheetCASH.cell(row, COL_KEY3)
            cellkey4 = sheetCASH.cell(row, COL_KEY4)
            if cellkey1.ctype == XL_CELL_EMPTY or cellkey1.ctype == XL_CELL_BLANK:
                return (row, COL_KEY1)
            if cellkey2.ctype == XL_CELL_EMPTY or cellkey2.ctype == XL_CELL_BLANK:
                return (row, COL_KEY2)
            if cellkey3.ctype == XL_CELL_EMPTY or cellkey3.ctype == XL_CELL_BLANK:
                return (row, COL_KEY3)
            if cellkey4.ctype == XL_CELL_EMPTY or cellkey4.ctype == XL_CELL_BLANK:
                return (row, COL_KEY4)
    #Se não for encontrado erro, retorna None.
    return None

def FindErrorDate():
    '''
    Pega cada workbook de solicitante, e vai selecionando as planilhas, que estão organizadas por
    mês do ano.
    Daí pega cada código e data de vencimento, que tem em cada linha, e compara com o mesmo em 
    CASH, na coluna deste solicitante.
    Se uma data+código não for achada: ERRO
    
    '''
    
    print 'COMPARANDO DATAS DA PLANILHA "'
    try:
        #Pegando a planilha de despesas detalhadas
        sheetCASH = BOOK_CASH.sheet_by_name(CASH_SHEET_DETAILED_EXPENSES)
    except:
        print "Erro: Não foi encontrada planilha ", CASH_SHEET_DETAILED_EXPENSES
        return 1
    
    #Lista de listas de tuplas do tipo: (<código>, <data de vencimento>).
    #A estrutura é: 
    #dada a lista reqinfolist cada item representa 1 solicitante. Este solicitante é respresentado
    #através de uma lista de tuplas de código + data_de_vencimento, que é uma chave única 
    #nas planilhas de CC (bradesco, itau, especie, ou outro) que serão abertas dinamicamente. 
    reqinfolist = []
    #A posição do solicitante na lista resquestor, é a mesma na lista reqinfolist
    #Inicializa a lista reqinfolist
    [reqinfolist.append([]) for i in range(len(REQUESTORS))]
    
    #Pegando o nome dos solicitantes (requestor) em Despesa de detalhada
    COL_REQ = CASH_SHEET_DETAILED_EXPENSES_COL_REQUESTOR
    COL_CODE = CASH_SHEET_DETAILED_EXPENSES_COL_CODE
    COL_DATE = CASH_SHEET_DETAILED_EXPENSES_COL_DUEDATE
    
    #Preencher a lista reqinfolist
    for row in range(sheetCASH.nrows):
        cellreq = sheetCASH.cell(row, COL_REQ)
        try:
            #Posição do solicitante: pos. Se não existir, gera ValueError.
            pos = REQUESTORS.index(cellreq.value)
        except:
            print 'Erro: lista dos solicitantes. Erro de software.'
            return 1
            
        #Pegar as informações do solicitante e colocar na lista reqinfolist
        #Da mesma linha que pegamos acima, pegar o código e a data de vencimento
        cellcode = sheetCASH.cell(row, COL_CODE)
        celldate = sheetCASH.cell(row, COL_DATE)
        #Com o código e a data de vencimento criar uma tupla, e colocar na posição "pos" de 
        #reqinfolist, no fim da lista desta posição
        reqinfolist[pos].append((cellcode, celldate))
    
    #Ordenar pela data
    for i in range(len(REQUESTORS)):
        l = reqinfolist[i] #Um solicitante por vez
        sorted(l, key=lambda f: f[1]) #Ordenar pelo segundo termo da tupla (a data)
    
    #Abrir cada arquivo de solicitante para comparar as datas com as de CASH
    #Considerando apenas as linhas de CASH que estão relacionadas à este solicitante
    for pos in range(len(REQUESTORS)):
        workbookCC = BOOK_CC_LIST[pos]
        
        #As planilhas de CC estão organizadas por mês
        for month in MONTH_ALIAS_LIST:
            #Dado o workbook, pegar a planilha do mês
            try:
                sheetCC = workbookCC.sheet_by_name(month)
            except:
                print "Erro: Não foi encontrada planilha ", month, ' do workbook: ', str(workbookCC)
                return 1
            
            #Comparação entre CASH e CC
            #Para cada linha da planilha sheetCC ver se o código e data de vencimento existem
            #na lista reqinfolist na posição (pos) do solicitante
            cellcode = None; celldate = None;
            for row in range(sheetCC.nrows):
                #Pegar o código da planilha do mês, em CC, e identificar se é um número
                cellcode = sheetCC.cell(row,CC_SHEET_COL_CODE)
                if not(cellcode.ctype == XL_CELL_NUMBER):
                    print 'Erro: não é um número a célula linha:', row, ' e coluna: ',\
                    CC_SHEET_COL_CODE, ' da planilha: ', month, ' do workbook: ', str(workbookCC)  
                    return 1
                
                #Pegar a data da planilha do mês, em CC, e identificar se é uma data
                celldate = sheetCC.cell(row,CC_SHEET_COL_DATE)
                if not(celldate.ctype == XL_CELL_DATE):
                    print 'Erro: não é uma data a célula linha:', row, ' e coluna: ',\
                    CC_SHEET_COL_CODE, ' da planilha: ', month, ' do workbook: ', str(workbookCC)  
                    return 1
                
                found = False
                for tupl in reqinfolist[pos]:
                    code = tupl[0]
                    date = tupl[1]
                    if code == cellcode.value and date == celldate.value:
                        found = True
                if not found:
                    print "Erro de comparação de código e data!"
                    print "   Não foi encontrado em CASH o código: ", code, " e data: ", date, "."
                    print "   Esse código (plano de contas) e data (data de vencimento) foram encontrados na célula de linha: ",\
                    row, " e coluna: ", month, " do workbook: ", str(workbookCC)
                    return 1
    return 0
    
def FindErrorYear(sheet, col):
    '''
    Descobre se o ano de vencimento de cada linha está errado
    
    '''
    
    for row in range(sheet.nrows):
        cell = sheet.cell(row, col)
        #Identificar se é uma data
        if not(cell.ctype == XL_CELL_DATE):
            #print 'Erro: não é uma data a célula linha:', row, ' e coluna: ',\
            #col, ' da planilha: ', month, ' do arquivo: ', filereq
            ccell = CCell(cell)
            ccell.setPos(row, col)  
            return ccell
        
        #structdate é a data estruturada no formato do python
        structdate = strptime(cell.value,"%d/%m/%Y")
        #Confere se o ano da data de vencimento de CC está certo (=YEAR)
        if not (structdate.tm_year == YEAR):
            ccell = CCell(cell)
            ccell.setPos(row, col)  
            return ccell
    
    return None
    
def FindErrorYearCC():
    #Para cada workbook
    for pos in range(len(REQUESTORS)):
        workbookCC = BOOK_CC_LIST[pos]
     
        #As planilhas de CC estão organizadas por mês
        for month in MONTH_ALIAS_LIST:
            #Dado o workbook, pegar a planilha do mês
            try:
                sheetCC = workbookCC.sheet_by_name(month)
            except:
                print "Erro: Não foi encontrada planilha ", month, ' do workbook: ', workbookCC
                return 1
     
            #O atributo structdate.tm_mon corresponde ao mês como um inteiro, pegando da lista
            #MONTH_ALIAS_LIST este inteiro decrementado, temos o nome deste mês no padrão das
            #planilhas
            #structdate_mon = MONTH_ALIAS_LIST[structdate.tm_mon - 1]
            col = CC_SHEET_COL_CODE
            
            #Chamar a função FindErrorYear para procurar erro na coluna col da planilha dada
            ccell = FindErrorYear(sheetCC, col)
            if not (ccell == None):
                print 'Erro: ano incorreto na célula linha:', ccell.getRow(), ' e coluna: ',\
                col, ' da planilha: ', month, ' do workbook: ', workbookCC  