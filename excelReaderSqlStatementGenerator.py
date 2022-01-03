import openpyxl
import os
import pandas as pd
import pprint
import collections
def excelDataGrabSQLGen(fileName, firstRow, lastRow, firstCol1, secondCol1):
    wb = openpyxl.load_workbook(fileName) #loads the excel file into the python IDE.
    dataSheet = input("What page of the file would you like to grab data from?") #Allows the user to selct a specific sheet within the excel file from which data can be taken.
    sheet = wb[dataSheet]
    sqlTableName = input("This version of the program is for generating sql statements, what is the table name?")
    sqlCol1 = input("What is the first sql column heading?")
    sqlCol2 = input("What is the second sql column heading?")
    #sqlCol3 = input("What is the third sql column heading?")
    for row in range(firstRow, lastRow + 1):
        firstValue = sheet.cell(row=row, column = firstCol1).value
        secondValue = sheet.cell(row = row, column = secondCol1).value
        #thirdValue = sheet.cell(row = row, column = thirdCol1).value

        print("INSERT INTO " + sqlTableName + "\n(" + sqlCol1 + ", " + sqlCol2 + ")\nVALUES( \'" + firstValue + "\', \'" + secondValue + "\');")




