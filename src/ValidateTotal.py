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
        print "Erro: Não foi encontrada planilha CONCILIAÇÃO"
        return 1
    
    #Imaginando que cada arquivo só contém os meses de 1 ano.
    #Esta repetição percorre cada sheet de CC e cada coluna de mês de CASH
    for i in range(0,12): 
        if (len(MONTH_ALIAS_LIST)-1) < i:
            return 1          
        #Pegando a coluna do mês de CASH. A partir da coluna C.
        colCASH = i + 1; #Para descontar a primeira coluna
        cellCASH = sheetCASH.cell(ROW_CASH, colCASH)
        #print 'Comparando planilha ', MONTH_ALIAS_LIST[i], ' com coluna ', colCASH
        #print 'Célula (', ROW_CASH + 1, ',', colCASH + 1, ') = ', cellCASH.value
        if not(cellCASH.ctype==XL_CELL_NUMBER):
            print "Erro na planilha de CASH: A célula da linha ", ROW_CASH, ' e coluna ', colCASH, ' não é um número.'
        
        #Pegando a sheet do mês de CC
        #A coluna é constante (G) e a linha é a última
        sheetCC = BOOK_CC.sheet_by_name(MONTH_ALIAS_LIST[i])
        #GetLastColCell retorna uma ComplCell
        ccellCC = GetLastColCell(sheetCC, COL_CC)
        rowCC = ccellCC.getRow(); colCC = ccellCC.getCol();
        #print 'Célula (', rowCC + 1, ',', colCC + 1, ') = ', ccellCC.getValue()
        if not(ccellCC.getType()==XL_CELL_NUMBER):
            print "Erro na planilha ", MONTH_ALIAS_LIST[i], " de CC: A célula da linha ", rowCC, ' e coluna ', colCC, ' não é um número.'
        
        if CompareMoney(ccellCC.getValue(), cellCASH.value):
            print "Total de ", MONTH_ALIAS_LIST[i], " ----------------------- OK"
        else:
            print "Total de ", MONTH_ALIAS_LIST[i], " ----------------------- ERRO"

def ValidateTotalIncome():
    print 'Validação do total da RECEITA'
    #Coluna da Receita na planilha de CC. Coluna G. É constantes na busca.
    COL_CC = 6
    #Linha dos totais na planilha de CASH. É constantes na busca. 
    #TODO: talvez fazer uma busca do "RECEITAS MÊS A MÊS"
    ROW_CASH = 4
    ValidateTotal(ROW_CASH, COL_CC)
    print 
#    #Valida o TOTAL das RECEITAS
#    print 'Comparando arquivos: ', FILE_CASH, ' e ', FILE_CC
#    print 'Validação do total da RECEITA'
#    sheetCASH = BOOK_CASH.sheet_by_name('RECEITAS')
#    
#    #Coluna da Receita na planilha de CC. Coluna G. É constantes na busca.
#    COL_INCOMINGS_CC = 6
#    #Linha dos totais na planilha de CASH. É constantes na busca.
#    ROW_CASH = sheetCASH.nrows-1
#    
#    #Imaginando que cada arquivo só contém os meses de 1 ano.
#    #Esta repetição percorre cada sheet de CC e cada coluna de mês de CASH
#    for i in range(0,12):
#        if (len(MONTH_ALIAS_LIST)-1) < i:
#            return          
#        #Pegando a coluna do mês de CASH. A partir da coluna C.
#        colCASH = 2 + i;
#        cellCASH = sheetCASH.cell(ROW_CASH, colCASH)
#        #print 'Comparando planilha ', MONTH_ALIAS_LIST[i], ' com coluna ', colCASH
#        #print 'Célula (', ROW_CASH + 1, ',', colCASH + 1, ') = ', cellCASH.value
#        if not(cellCASH.ctype==XL_CELL_NUMBER):
#            print "Erro na planilha de CASH: A célula da linha ", ROW_CASH, ' e coluna ', colCASH, ' não é um número.'
#        
#        #Pegando a sheet do mês de CC
#        #A coluna é constante (G) e a linha é a última
#        sheetCC = BOOK_CC.sheet_by_name(MONTH_ALIAS_LIST[i])
#        #GetLastColCell retorna uma ComplCell
#        ccellCC = GetLastColCell(sheetCC, COL_INCOMINGS_CC)
#        rowCC = ccellCC.getRow(); colCC = ccellCC.getCol();
#        #print 'Célula (', rowCC + 1, ',', colCC + 1, ') = ', ccellCC.getValue()
#        if not(ccellCC.getType()==XL_CELL_NUMBER):
#            print "Erro na planilha ", MONTH_ALIAS_LIST[i], " de CC: A célula da linha ", rowCC, ' e coluna ', colCC, ' não é um número.'
#        
#        if CompareMoney(ccellCC.getValue(), cellCASH.value):
#            print "Total Receita de ", MONTH_ALIAS_LIST[i], " ----------------------- OK"
#        else:
#            print "Total Receita de ", MONTH_ALIAS_LIST[i], " ----------------------- ERRO"

def ValidateTotalExpense():
    print 'Validação do total da DESPESA'
    #Coluna da Receita na planilha de CC. Coluna G. É constantes na busca.
    COL_CC = 5
    #Linha dos totais na planilha de CASH. É constantes na busca. 
    #TODO: talvez fazer uma busca do "DESPESAS MÊS A MÊS"
    ROW_CASH = 6
    ValidateTotal(ROW_CASH, COL_CC)
    print 

def ValidateTotalBalance():
    print 'Validação do total do SALDO'
    #Coluna da Receita na planilha de CC. Coluna G. É constantes na busca.
    COL_CC = 4
    #Linha dos totais na planilha de CASH. É constantes na busca. 
    #TODO: talvez fazer uma busca do "SALDO CONTA CORRENTE"
    ROW_CASH = 29
    ValidateTotal(ROW_CASH, COL_CC)
    print
