#!/usr/bin/env python3
# -*- coding: utf8 -*-

# Модуль распознавания файла от "Автомеханика"

import openpyxl

import pymongo

class TAutomexanika:
    # Загрузка ексель файла фирмы Автомеханика в базу Mongo

    def __init__(self, xls_file):
        """Constructor"""
        self.xls_file = xls_file

    def xlsToMongo(self):
        xls_document = openpyxl.load_workbook(self.xls_file)
        xls_worksheet = xls_document.worksheets[0] # Назначаем индекс листа, который будит парсится

        mongo_db = pymongo.MongoClient()["aridan"]
        mongo_db_collection = mongo_db["all"]

        shop = "automexanika" # Название фирмы, от которой пришёл прайс-лист
        level = ""
        group = ""

        column_count = xls_worksheet.max_column
        row_count    = xls_worksheet.max_row

        for row in range(1, row_count):
            index  = 0

            record = ""

            code = ""
            type_product = ""
            brend = ""
            oem_char = ""
            oem_number = ""
            oem_all_number = ""
            analog = ""
            name = ""
            new_product = ""
            price = ""
            avaible = ""
            incoming = ""

            for column in range(1, column_count):
                index += 1
                value = str(xls_worksheet.cell(row, column).value).strip()

                if (index == 1): # Пытаемся получить уровень
                    level_var = value

                if (index == 2): # Пытаемся полчить группу
                    if (value != "0" and value != "None"):
                        if (level_var == value):
                            level = level_var # Точнополучили уровень
                            #print(level+":")
                        else:
                            group = value # Точно получили группу
                            #print("    " + group)
                        continue # Пескакиваем на следующую итерацию, так как ловить болше нечего

                if (index == 3): # Пытаемся получить код
                    code = value

                if (index == 4): # Пытаемся получить тип
                    type_product = value

                if (index == 5): # Пытаемся получить производителя
                    brend = value

                if (index == 6): # Пытаемся получить инициалы детали у производителя
                    oem_char = value

                if (index == 7): # Пытаемся получить номер детали у производителя
                    oem_number = value

                if (index == 8): # Пытаемся получить все номера детали у производителя
                    oem_all_number = value

                if (index == 9): # Пытаемся получить номера совместимых деталей
                    analog = value

                if (index == 10): # Пытаемся получить наименование
                    name = value

                if (index == 11): # Пытаемся получить я вляется ли товар новинкой
                    new_product = value

                if (index == 12): # Пытаемся получить цену
                    price = value

                if (index == 13): # Пытаемся получить наличие товара на складе
                    avaible = value

                if (index == 14): # Пытаемся получить приход товара на складе
                    incoming = value

            if (code != "None"):
                record = "          " +\
                         code +" | "+ \
                         type_product +" | "+ \
                         brend +" | "+ \
                         oem_char +" | "+ \
                         oem_number +" | "+ \
                         oem_all_number +" | "+ \
                         analog +" | "+ \
                         name +" | "+ \
                         new_product +" | "+ \
                         price +" | "+ \
                         avaible +" | "+ \
                         incoming +"|"
                #print (record)

                mongo_doc = {"shop":shop,
                             "level":level,
                             "group":group,
                             "code":code,
                             "type_product":type_product,
                             "brend":brend,
                             "oem_char":oem_char,
                             "oem_number":oem_number,
                             "oem_all_number":oem_all_number,
                             "analog":analog,
                             "name":name,
                             "new_product":new_product,
                             "price":price,
                             "avaible":avaible,
                             "incoming":incoming}

                mongo_db_collection.insert_one(mongo_doc).inserted_id
#--------------------------------------------------------------------------------------------------------------






