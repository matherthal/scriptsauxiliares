#!/usr/bin/python
# -*- coding: iso-8859-1 -*-

#Esta classe contém não somente o valor e tipo da célula, como sua localização
class CCell:
    def __init__(self, cell):
        self._value = cell.value
        self._ctype = cell.ctype
        self._row = None
        self._col = None
    
    def setPos(self, row, col):
        self._row = row
        self._col = col
        
    def getPos(self):
        return (self._row,self._col)
    
    def setRow(self, row):
        self._row = row
        
    def getRow(self):
        return self._row
    
    def setCol(self, col):
        self._col = col
        
    def getCol(self):
        return self._col

    def setValue(self, value):
        self._value = value
        
    def getValue(self):
        return self._value
        
    def setType(self, type):
        self._ctype = type
        
    def getType(self):
        return self._ctype