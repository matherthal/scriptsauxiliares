#!/usr/bin/python
# -*- coding: iso-8859-1 -*-
#import os, sys

from ValidateTotal import ValidateTotalIncome, ValidateTotalExpense, ValidateTotalBalance
from GUI import FileDialog
from ExcelAuxFuncs import SetBookCASH, SetBookCC
import Tkinter

def StartWorkBooks():
    root = Tkinter.Tk()
    FileDialog(root).pack()
    root.mainloop()
    dialog = FileDialog()
    filepath = dialog.GetFileName('Informe o arquivo de CASH')
    SetBookCASH(filepath)
    filepath = dialog.GetFileName('Informe o arquivo de CC')
    SetBookCC(filepath)

def main():
    print "-----------------------------------------------------"
    print "|                                                   |"
    print "|                 ESQUILO VALIDADOR                 |"
    print "|                                                   |"
    print "|                              Matheus de Sá Erthal |"
    print "-----------------------------------------------------"
    print 
    StartWorkBooks()
    ValidateTotalIncome()
    ValidateTotalExpense()
    ValidateTotalBalance()
        
#import msvcrt
#msvcrt.getch()

if __name__ == "__main__":
    main()