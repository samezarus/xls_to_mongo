#!/usr/bin/env python3
# -*- coding: utf8 -*-

import openpyxl

#import pymongo

import automexanika
 
file = "/home/sameza/1/automexanika.xlsx"

#Automexanika = automexanika.TAutomexanika(file)
#Automexanika.xlsToMongo()

file = "/home/sameza/1/autovladzapchast.xlsx"

xls_document = openpyxl.load_workbook(file)
xls_worksheet = xls_document.worksheets[0] # Назначаем индекс листа, который будит парсится

def getColor(xls_cell):
    result = ""

    Colors = openpyxl.styles.colors.COLOR_INDEX
    i = xls_cell.fill.start_color.index
    result = str(Colors[i])

    return result


xls_cell = xls_worksheet.cell(9, 2)

print (getColor(xls_cell))