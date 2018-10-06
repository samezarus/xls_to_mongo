#!/usr/bin/env python3
# -*- coding: utf8 -*-

import openpyxl

import pymongo

from bson import json_util # Для записи строки в Mongo (интегрирован в pymongo)

import automexanika
import autovladzapchast

# --------------------------------------------------------------------------------------------

mongo_db = pymongo.MongoClient()["aridan"]
mongo_db_collection = mongo_db["all"]
mongo_db_collection.remove({})

file = "/home/sameza/1/automexanika.xlsx"

Automexanika = automexanika.TAutomexanika(file)
Automexanika.mongo_db_name = "aridan"
Automexanika.mongo_db_collection_name = "all"
#Automexanika.xlsToMongo()

file = "/home/sameza/1/autovladzapchast.xlsx"

AutoVladZapchast = autovladzapchast.TAutoVladZapchast(file)
AutoVladZapchast.mongo_db_name = "aridan"
AutoVladZapchast.mongo_db_collection_name = "all"
AutoVladZapchast.xlsToMongo()





#xls_document  = openpyxl.load_workbook(file)
#xls_worksheet = xls_document.worksheets[0] # Назначаем индекс листа, который будит парсится

#xls_column_count = xls_worksheet.max_column
#xls_row_count    = xls_worksheet.max_row

#for row in range(2, xls_row_count):
#    record = "| "
    #for column in range(1, xls_column_count):
        #xls_cell = xls_worksheet.cell(row, column)
        #print (getColor(xls_cell))
        #value = str(xls_cell.value).strip()
        #record = record + value + " | "
    #print (record.split("\n"))
    #print ("---------------------------------------")

#    xls_cell = xls_worksheet.cell(row, 2)
#    #print(str(row) + " - " + getColor(xls_cell))
#    if (getColor(xls_cell) == "8"):
#        print (xls_cell.value)

#    if (getColor(xls_cell) == "22"):
#        print ("  "+xls_cell.value)



