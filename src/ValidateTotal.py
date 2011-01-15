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
        print "Erro: N�o foi encontrada planilha ", CASH_SHEET_DETAILED_EXPENSES
        return 1
    
    #Pegando o nome dos solicitantes (requestor) em Despesa de detalhada
    COL_REQ = CASH_SHEET_DETAILED_EXPENSES_COL_REQUESTOR
    
    #Preencher as listas REQUESTORS
    for row in range(sheetCASH.nrows):
        #Se o solicitante n�o estiver na lista, inseri-lo
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
                print "Erro: N�o foi encontrado o arquivo ", filereq
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
        print "Erro: N�o foi encontrada planilha CONCILIA��O"
        return 1
    
    #Imaginando que cada arquivo s� cont�m os meses de 1 ano.
    #Esta repeti��o percorre cada sheet de CC e cada coluna de m�s de CASH
    for i in range(0,12): 
        if (len(MONTH_ALIAS_LIST)-1) < i:
            return 1          
        #Pegando a coluna do m�s de CASH. A partir da coluna C.
        colCASH = i + 1; #Para descontar a primeira coluna
        val = 0
        for j in range(0,len(ROWS_CASH)):
            cellCASH = sheetCASH.cell(ROWS_CASH[j], colCASH)
            #print 'Comparando planilha ', MONTH_ALIAS_LIST[i], ' com coluna ', colCASH
            #print 'C�lula (', ROW_CASH + 1, ',', colCASH + 1, ') = ', cellCASH.value
            if cellCASH.ctype==XL_CELL_NUMBER:
                val += cellCASH.value
            #else:
            #    print "Erro na planilha de CASH: A c�lula da linha ", ROWS_CASH[j], ' e coluna ', colCASH, ' n�o � um n�mero.'
            #DEBUG print ROWS_CASH[j], ',', colCASH, ': ', str(cellCASH.value), ' tipo: ', cellCASH.ctype 
        
        #Pegando a sheet do m�s de CC especie
        #A coluna � constante (G) e a linha � a �ltima
        sheetCC = BOOK_MONEY.sheet_by_name(MONTH_ALIAS_LIST[i])
        #GetLastColCell retorna uma CCell
        ccellCC = GetLastColCell(sheetCC, COL_CC)
        rowCC = ccellCC.getRow(); colCC = ccellCC.getCol();
        #print 'C�lula (', rowCC + 1, ',', colCC + 1, ') = ', ccellCC.getValue()
        if not(ccellCC.getType()==XL_CELL_NUMBER):
            print "Erro na planilha ", MONTH_ALIAS_LIST[i], " de CC: A c�lula da linha ", rowCC, ' e coluna ', colCC, ' n�o � um n�mero.'
        
        if CompareMoney(ccellCC.getValue(), val):
            print "Total de ", MONTH_ALIAS_LIST[i], " ----------------------- OK"
        else:
            print "Total de ", MONTH_ALIAS_LIST[i], " ----------------------- ERRO"
    return 0

def ValidateTotalIncome():
    '''
    Valida o TOTAL da Receita
    
    '''
    print '1. VALIDA��O DO TOTAL DA RECEITA'
    #Coluna da Receita na planilha de CC. Coluna J. � constantes na busca.
    COL_CC = 6
    #COL_CC = 9
    #Linhas importantes para c�lculo da receita na planilha de CASH. � constantes na busca. 
    ROWS_CASH = [4,17,18,19,20]
    ValidateTotal(ROWS_CASH, COL_CC)
    print 

def ValidateTotalExpense():
    '''
    Valida o TOTAL da Despesa
    
    '''
    
    print '2. VALIDA��O DO TOTAL DA DESPESA'
    #Coluna da Despesa na planilha de CC. Coluna I. � constantes na busca.
    COL_CC = 5
    #COL_CC = 8
    #Linhas importantes para c�lculo da despesa na planilha de CASH. � constantes na busca. 
    ROWS_CASH = [6,24,25,26,27,28,29]
    ValidateTotal(ROWS_CASH, COL_CC)
    print 

def ValidateTotalBalance():
    '''
    Valida o TOTAL do Saldo
    
    '''
    
    print '3. VALIDA��O DO SALDO'
    #Coluna do Saldo na planilha de CC. Coluna K. � constantes na busca.
    COL_CC = 7
    #COL_CC = 10
    #Linha do Saldo 
    ROWS_CASH = [31]
    ValidateTotal(ROWS_CASH, COL_CC)
    print

def FindErrorConstructionKeyExistence():
    '''
    Busca erro de n�o exist�ncia de chave para um c�digo. Na planilha Despesas Detalhadas, arquivo CASH 
    Retorna uma tupla correspondente � posi��o da c�lula de erro. Sen�o retorna None.
    
    '''
    
    print 'BUSCANDO ERRO DE CHAVE NA PLANILHA "', CASH_SHEET_DETAILED_EXPENSES, '" em CASH.'
    try:
        sheetCASH = BOOK_CASH.sheet_by_name(CASH_SHEET_DETAILED_EXPENSES)
    except:
        print "Erro: N�o foi encontrada planilha ", CASH_SHEET_DETAILED_EXPENSES
        return 1
    
    #Coluna "C�digo da Raz�o Geral"
    COL_CODE = CASH_SHEET_DETAILED_EXPENSES_COL_CODE
    #Colunas de chaves
    COL_KEY1 = CASH_SHEET_DETAILED_EXPENSES_COL_KEY1
    COL_KEY2 = CASH_SHEET_DETAILED_EXPENSES_COL_KEY2
    COL_KEY3 = CASH_SHEET_DETAILED_EXPENSES_COL_KEY3
    COL_KEY4 = CASH_SHEET_DETAILED_EXPENSES_COL_KEY4
    
    for row in range(sheetCASH.nrows):
        cellcode = sheetCASH.cell(row, COL_CODE)
        #Se existir o c�digo da coluna A, e n�o existir uma das chaves das �ltimas colunas: Erro
        if cellcode.ctype == XL_CELL_NUMBER:
            #Pesquisando nas 4 chaves da linha. Se alguma estiver em branco ou vazia, returna sua posi��o
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
    #Se n�o for encontrado erro, retorna None.
    return None

def FindErrorDate():
    '''
    Pega cada workbook de solicitante, e vai selecionando as planilhas, que est�o organizadas por
    m�s do ano.
    Da� pega cada c�digo e data de vencimento, que tem em cada linha, e compara com o mesmo em 
    CASH, na coluna deste solicitante.
    Se uma data+c�digo n�o for achada: ERRO
    
    '''
    
    print 'COMPARANDO DATAS DA PLANILHA "'
    try:
        #Pegando a planilha de despesas detalhadas
        sheetCASH = BOOK_CASH.sheet_by_name(CASH_SHEET_DETAILED_EXPENSES)
    except:
        print "Erro: N�o foi encontrada planilha ", CASH_SHEET_DETAILED_EXPENSES
        return 1
    
    #Lista de listas de tuplas do tipo: (<c�digo>, <data de vencimento>).
    #A estrutura �: 
    #dada a lista reqinfolist cada item representa 1 solicitante. Este solicitante � respresentado
    #atrav�s de uma lista de tuplas de c�digo + data_de_vencimento, que � uma chave �nica 
    #nas planilhas de CC (bradesco, itau, especie, ou outro) que ser�o abertas dinamicamente. 
    reqinfolist = []
    #A posi��o do solicitante na lista resquestor, � a mesma na lista reqinfolist
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
            #Posi��o do solicitante: pos. Se n�o existir, gera ValueError.
            pos = REQUESTORS.index(cellreq.value)
        except:
            print 'Erro: lista dos solicitantes. Erro de software.'
            return 1
            
        #Pegar as informa��es do solicitante e colocar na lista reqinfolist
        #Da mesma linha que pegamos acima, pegar o c�digo e a data de vencimento
        cellcode = sheetCASH.cell(row, COL_CODE)
        celldate = sheetCASH.cell(row, COL_DATE)
        #Com o c�digo e a data de vencimento criar uma tupla, e colocar na posi��o "pos" de 
        #reqinfolist, no fim da lista desta posi��o
        reqinfolist[pos].append((cellcode, celldate))
    
    #Ordenar pela data
    for i in range(len(REQUESTORS)):
        l = reqinfolist[i] #Um solicitante por vez
        sorted(l, key=lambda f: f[1]) #Ordenar pelo segundo termo da tupla (a data)
    
    #Abrir cada arquivo de solicitante para comparar as datas com as de CASH
    #Considerando apenas as linhas de CASH que est�o relacionadas � este solicitante
    for pos in range(len(REQUESTORS)):
        workbookCC = BOOK_CC_LIST[pos]
        
        #As planilhas de CC est�o organizadas por m�s
        for month in MONTH_ALIAS_LIST:
            #Dado o workbook, pegar a planilha do m�s
            try:
                sheetCC = workbookCC.sheet_by_name(month)
            except:
                print "Erro: N�o foi encontrada planilha ", month, ' do workbook: ', str(workbookCC)
                return 1
            
            #Compara��o entre CASH e CC
            #Para cada linha da planilha sheetCC ver se o c�digo e data de vencimento existem
            #na lista reqinfolist na posi��o (pos) do solicitante
            cellcode = None; celldate = None;
            for row in range(sheetCC.nrows):
                #Pegar o c�digo da planilha do m�s, em CC, e identificar se � um n�mero
                cellcode = sheetCC.cell(row,CC_SHEET_COL_CODE)
                if not(cellcode.ctype == XL_CELL_NUMBER):
                    print 'Erro: n�o � um n�mero a c�lula linha:', row, ' e coluna: ',\
                    CC_SHEET_COL_CODE, ' da planilha: ', month, ' do workbook: ', str(workbookCC)  
                    return 1
                
                #Pegar a data da planilha do m�s, em CC, e identificar se � uma data
                celldate = sheetCC.cell(row,CC_SHEET_COL_DATE)
                if not(celldate.ctype == XL_CELL_DATE):
                    print 'Erro: n�o � uma data a c�lula linha:', row, ' e coluna: ',\
                    CC_SHEET_COL_CODE, ' da planilha: ', month, ' do workbook: ', str(workbookCC)  
                    return 1
                
                found = False
                for tupl in reqinfolist[pos]:
                    code = tupl[0]
                    date = tupl[1]
                    if code == cellcode.value and date == celldate.value:
                        found = True
                if not found:
                    print "Erro de compara��o de c�digo e data!"
                    print "   N�o foi encontrado em CASH o c�digo: ", code, " e data: ", date, "."
                    print "   Esse c�digo (plano de contas) e data (data de vencimento) foram encontrados na c�lula de linha: ",\
                    row, " e coluna: ", month, " do workbook: ", str(workbookCC)
                    return 1
    return 0
    
def FindErrorYear(sheet, col):
    '''
    Descobre se o ano de vencimento de cada linha est� errado
    
    '''
    
    for row in range(sheet.nrows):
        cell = sheet.cell(row, col)
        #Identificar se � uma data
        if not(cell.ctype == XL_CELL_DATE):
            #print 'Erro: n�o � uma data a c�lula linha:', row, ' e coluna: ',\
            #col, ' da planilha: ', month, ' do arquivo: ', filereq
            ccell = CCell(cell)
            ccell.setPos(row, col)  
            return ccell
        
        #structdate � a data estruturada no formato do python
        structdate = strptime(cell.value,"%d/%m/%Y")
        #Confere se o ano da data de vencimento de CC est� certo (=YEAR)
        if not (structdate.tm_year == YEAR):
            ccell = CCell(cell)
            ccell.setPos(row, col)  
            return ccell
    
    return None
    
def FindErrorYearCC():
    #Para cada workbook
    for pos in range(len(REQUESTORS)):
        workbookCC = BOOK_CC_LIST[pos]
     
        #As planilhas de CC est�o organizadas por m�s
        for month in MONTH_ALIAS_LIST:
            #Dado o workbook, pegar a planilha do m�s
            try:
                sheetCC = workbookCC.sheet_by_name(month)
            except:
                print "Erro: N�o foi encontrada planilha ", month, ' do workbook: ', workbookCC
                return 1
     
            #O atributo structdate.tm_mon corresponde ao m�s como um inteiro, pegando da lista
            #MONTH_ALIAS_LIST este inteiro decrementado, temos o nome deste m�s no padr�o das
            #planilhas
            #structdate_mon = MONTH_ALIAS_LIST[structdate.tm_mon - 1]
            col = CC_SHEET_COL_CODE
            
            #Chamar a fun��o FindErrorYear para procurar erro na coluna col da planilha dada
            ccell = FindErrorYear(sheetCC, col)
            if not (ccell == None):
                print 'Erro: ano incorreto na c�lula linha:', ccell.getRow(), ' e coluna: ',\
                col, ' da planilha: ', month, ' do workbook: ', workbookCC  