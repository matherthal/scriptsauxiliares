#!/usr/bin/python
# -*- coding: iso-8859-1 -*-

from xlrd import XL_CELL_NUMBER
from CCell import ComplCell
from ExcelAuxFuncs import BOOK_CASH, BOOK_CC, FILE_CASH, FILE_CC, MONTH_ALIAS_LIST, GetLastColCell, CompareMoney , GetRowIndexfromCol, CASH_SHEET

def ValidateTotal(ROW_CASH, COL_CC):
    #Valida o TOTAL da Receita, Despesa ou Saldo
    print 'Comparando arquivos: ', FILE_CASH, ' e ', FILE_CC
    try:
        sheetCASH = BOOK_CASH.sheet_by_name(CASH_SHEET)
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
        cellCASH = sheetCASH.cell(ROW_CASH, colCASH)
        #print 'Comparando planilha ', MONTH_ALIAS_LIST[i], ' com coluna ', colCASH
        #print 'C�lula (', ROW_CASH + 1, ',', colCASH + 1, ') = ', cellCASH.value
        if not(cellCASH.ctype==XL_CELL_NUMBER):
            print "Erro na planilha de CASH: A c�lula da linha ", ROW_CASH, ' e coluna ', colCASH, ' n�o � um n�mero.'
        
        #Pegando a sheet do m�s de CC
        #A coluna � constante (G) e a linha � a �ltima
        sheetCC = BOOK_CC.sheet_by_name(MONTH_ALIAS_LIST[i])
        #GetLastColCell retorna uma ComplCell
        ccellCC = GetLastColCell(sheetCC, COL_CC)
        rowCC = ccellCC.getRow(); colCC = ccellCC.getCol();
        #print 'C�lula (', rowCC + 1, ',', colCC + 1, ') = ', ccellCC.getValue()
        if not(ccellCC.getType()==XL_CELL_NUMBER):
            print "Erro na planilha ", MONTH_ALIAS_LIST[i], " de CC: A c�lula da linha ", rowCC, ' e coluna ', colCC, ' n�o � um n�mero.'
        
        if CompareMoney(ccellCC.getValue(), cellCASH.value):
            print "Total de ", MONTH_ALIAS_LIST[i], " ----------------------- OK"
        else:
            print "Total de ", MONTH_ALIAS_LIST[i], " ----------------------- ERRO"

def ValidateTotalIncome():
    print 'Valida��o do total da RECEITA'
    #Coluna da Receita na planilha de CC. Coluna G. � constantes na busca.
    COL_CC = 6
    #Linha dos totais na planilha de CASH. � constantes na busca. 
    #TODO: talvez fazer uma busca do "RECEITAS M�S A M�S"
    ROW_CASH = 4
    ValidateTotal(ROW_CASH, COL_CC)
    print 
#    #Valida o TOTAL das RECEITAS
#    print 'Comparando arquivos: ', FILE_CASH, ' e ', FILE_CC
#    print 'Valida��o do total da RECEITA'
#    sheetCASH = BOOK_CASH.sheet_by_name('RECEITAS')
#    
#    #Coluna da Receita na planilha de CC. Coluna G. � constantes na busca.
#    COL_INCOMINGS_CC = 6
#    #Linha dos totais na planilha de CASH. � constantes na busca.
#    ROW_CASH = sheetCASH.nrows-1
#    
#    #Imaginando que cada arquivo s� cont�m os meses de 1 ano.
#    #Esta repeti��o percorre cada sheet de CC e cada coluna de m�s de CASH
#    for i in range(0,12):
#        if (len(MONTH_ALIAS_LIST)-1) < i:
#            return          
#        #Pegando a coluna do m�s de CASH. A partir da coluna C.
#        colCASH = 2 + i;
#        cellCASH = sheetCASH.cell(ROW_CASH, colCASH)
#        #print 'Comparando planilha ', MONTH_ALIAS_LIST[i], ' com coluna ', colCASH
#        #print 'C�lula (', ROW_CASH + 1, ',', colCASH + 1, ') = ', cellCASH.value
#        if not(cellCASH.ctype==XL_CELL_NUMBER):
#            print "Erro na planilha de CASH: A c�lula da linha ", ROW_CASH, ' e coluna ', colCASH, ' n�o � um n�mero.'
#        
#        #Pegando a sheet do m�s de CC
#        #A coluna � constante (G) e a linha � a �ltima
#        sheetCC = BOOK_CC.sheet_by_name(MONTH_ALIAS_LIST[i])
#        #GetLastColCell retorna uma ComplCell
#        ccellCC = GetLastColCell(sheetCC, COL_INCOMINGS_CC)
#        rowCC = ccellCC.getRow(); colCC = ccellCC.getCol();
#        #print 'C�lula (', rowCC + 1, ',', colCC + 1, ') = ', ccellCC.getValue()
#        if not(ccellCC.getType()==XL_CELL_NUMBER):
#            print "Erro na planilha ", MONTH_ALIAS_LIST[i], " de CC: A c�lula da linha ", rowCC, ' e coluna ', colCC, ' n�o � um n�mero.'
#        
#        if CompareMoney(ccellCC.getValue(), cellCASH.value):
#            print "Total Receita de ", MONTH_ALIAS_LIST[i], " ----------------------- OK"
#        else:
#            print "Total Receita de ", MONTH_ALIAS_LIST[i], " ----------------------- ERRO"

def ValidateTotalExpense():
    print 'Valida��o do total da DESPESA'
    #Coluna da Receita na planilha de CC. Coluna G. � constantes na busca.
    COL_CC = 5
    #Linha dos totais na planilha de CASH. � constantes na busca. 
    #TODO: talvez fazer uma busca do "DESPESAS M�S A M�S"
    ROW_CASH = 6
    ValidateTotal(ROW_CASH, COL_CC)
    print 

def ValidateTotalBalance():
    print 'Valida��o do total do SALDO'
    #Coluna da Receita na planilha de CC. Coluna G. � constantes na busca.
    COL_CC = 4
    #Linha dos totais na planilha de CASH. � constantes na busca. 
    #TODO: talvez fazer uma busca do "SALDO CONTA CORRENTE"
    ROW_CASH = 29
    ValidateTotal(ROW_CASH, COL_CC)
    print
