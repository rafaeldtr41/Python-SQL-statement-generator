import openpyxl
import os
import pandas as pd
import pprint
import collections




def generate_half(*val, nomb):
    #This generate the half of the insert
    value = "Insert into " + nomb + " (\n"
    for x in val:
        value  = value + x + ",\n"
    value = value + ")\n values (\n"
    return value
        

def excelDataGrabSQLGen(fileName, firstRow, lastRow, firstCol1, secondCol1):
    wb = openpyxl.load_workbook(fileName) #loads the excel file into the python IDE.
    dataSheet = input("What page of the file would you like to grab data from?") #Allows the user to selct a specific sheet within the excel file from which data can be taken.
    sheet = wb[dataSheet]
    sqlTableName = input("This version of the program is for generating sql statements, what is the table name?")
    handler = True
    lista = ()
    while handler:
        sqlCol = input("What is the sql column heading?, type 0 when you done")
        if sqlCol != "0": 
            lista.__add__(sqlCol)
        else:
            handler = False

    aux = generate_half(lista)

    """
    sqlCol1 = input("What is the first sql column heading?")
    sqlCol2 = input("What is the second sql column heading?")
    #sqlCol3 = input("What is the third sql column heading?")
    """
    for row in range(firstRow, lastRow + 1):
        aux1 = aux #To dont charge the same function n times
        for cell in row:

            if lista.__contains__(cell.column):
                aux1 = aux1 + cell.value + ",\n"
        aux1 = aux1 + "\n);"
    print(aux1)




