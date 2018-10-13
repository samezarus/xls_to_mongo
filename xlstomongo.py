#!/usr/bin/env python3
# -*- coding: utf8 -*-

# Модуль распознавания файла от "Автомеханика"

import os

import openpyxl
import xlrd

import pymongo

from bson import json_util # Для записи строки в Mongo (интегрирован в pymongo)

import re

class TXlsToMongo:
    # Загрузка ексель файла фирмы Автомеханика в базу Mongo

    mongo_db_name = ''
    mongo_db_collection_name = ''
    shop = ''  # Название фирмы, от которой пришёл прайс-лист
    xls_file = '' # эксель файл
    xls_worksheet_namber = 0 # Номер вкладки в эксель файле

    def __init__(self):
        """Constructor"""
        xyz = ''
    #
    def jsonToMongo(self, json_string):
        return json_util.loads(json_string)
    #
    def jsonReplase(self,string):
        if (len(string) > 0):
            # Замена символов учавствующих в синтаксисе json
            string = re.sub('"', '″', string)
            # Удаляем литеры из строки
            string = re.sub('\a|\b|\f|\r|\t|\v|', '', string)

            # если в строке есть литер (переноса строки \n), каждую новую строку обрамляем двойными ковычками и разделяем запятыми
            string = re.sub('\n', '","', string)
            string = string.replace('\\', '/')  # Замена символа "\" на "/"
            if (string.find('"') != -1):
                string = '"' + string + '"'
            #
        return string

    #
    def xlsToMongo(self):
        mongo_db            = pymongo.MongoClient()[self.mongo_db_name]
        mongo_db_collection = mongo_db[self.mongo_db_collection_name]

        extension = os.path.splitext(self.xls_file)[1]
        if (extension == '.xlsx'):
            xls_document     = openpyxl.load_workbook(self.xls_file)
            xls_worksheet    = xls_document.worksheets[self.xls_worksheet_namber]  # Назначаем индекс листа, который будит парсится
            xls_column_count = xls_worksheet.max_column
            xls_row_count    = xls_worksheet.max_row
        else:
            xls_document     = xlrd.open_workbook(self.xls_file)
            xls_worksheet    = xls_document.sheet_by_index(self.xls_worksheet_namber)
            xls_column_count = xls_worksheet.ncols
            xls_row_count    = xls_worksheet.nrows


        mongo_db_collection.remove({}) # Очищаем коллекцию(таблицу)

        for row in range(1, xls_row_count):
            json_doc = ''
            for column in range(1, xls_column_count):
                field = 'Fileld' + str(column)
                xls_cell = xls_worksheet.cell(row, column)
                value = str(xls_cell.value).strip()

                if (len(value) > 0):

                    value = self.jsonReplase(value)

                    if (value[0] == '"' and value[len(value) - 1] == '"'):
                        value = '[' + value + ']'
                    else:
                        value = '"' + value + '"'
                else:
                    value = '""'
                json_doc = json_doc + '"' + field + '":' + value + ','

            json_doc = json_doc[0:len(json_doc) - 1]
            json_doc = r'{' + json_doc + '}'
            print(json_doc)
            mongo_doc = self.jsonToMongo(json_doc)
            mongo_db_collection.insert_one(mongo_doc)
#--------------------------------------------------------------------------------------------------------------






