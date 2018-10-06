#!/usr/bin/env python3
# -*- coding: utf8 -*-

import openpyxl

import pymongo

from bson import json_util # Для записи строки в Mongo (интегрирован в pymongo)

import automexanika

# --------------------------------------------------------------------------------------------
def getColor(xls_cell):
    result = ""

    Colors = openpyxl.styles.colors.COLOR_INDEX
    i = xls_cell.fill.start_color.index
    result = str(Colors[i])

    return result

# --------------------------------------------------------------------------------------------
mongo_db = pymongo.MongoClient()["aridan"]
mongo_db_collection = mongo_db["all"]
mongo_db_collection.remove({})

file = "/home/sameza/1/automexanika.xlsx"

Automexanika = automexanika.TAutomexanika(file)
Automexanika.mongo_db_name = "aridan"
Automexanika.mongo_db_collection_name = "all"
Automexanika.xlsToMongo()

file = "/home/sameza/1/autovladzapchast.xlsx"

#xls_document = openpyxl.load_workbook(file)
#xls_worksheet = xls_document.worksheets[0] # Назначаем индекс листа, который будит парсится

#xls_column_count = xls_worksheet.max_column
#xls_row_count = xls_worksheet.max_row

#for row in range(1, xls_row_count):
#    record = "| "
#    for column in range(1, xls_column_count):
#        value = str(xls_worksheet.cell(row, column).value).strip()
#        record = record + value + " | "
#    print (record.split("\n"))
#    print ("---------------------------------------")


# Запись в Mongo из строки
#mongo_db = pymongo.MongoClient()["aridan"]
#mongo_db_collection = mongo_db["all"]
#json_body = '{"test":"ok", "value":1}'
#data = json_util.loads(json_body)
#mongo_db_collection.insert_one(data)

