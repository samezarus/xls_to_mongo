#!/usr/bin/env python3
# -*- coding: utf8 -*-

import os
import re
# --------------------------------------------------------------------------------------------
class TCommandParam:
    command = ''
    param = ''
#
class TCreateItem:
    newFieldName  = ''
    newFieldValue = ''
    fromFields    = []
    mathCommand   = ''
    type = ''
#
def getConfigCreate(createNewFieldCommands, recordsList):
    #
    result = TCreateItem()

    for i in range(0, len(createNewFieldCommands)):
        if createNewFieldCommands[i] == ':':
            print (i)

    return result
#
def getCommandParam(commandParamString):
    CommandParam = TCommandParam()
    p = commandParamString.find(' ')
    if p != -1:
        CommandParam.command = commandParamString[1:p]
        CommandParam.param = commandParamString[p +1:len(commandParamString)]

    return CommandParam
#
def getFrom(fromString):
    result = []
    fromString = fromString + ' '

    p = 0
    for i in range(0, len(fromString)):
        if fromString[i] == ' ':
            result.append(fromString[p:i])
            p = i +1

    return result
#
def getMath(mathString):
    CommandParam = TCommandParam()

    CommandParam.command = mathString[0]
    CommandParam.param = mathString[1:len(mathString)]

    return CommandParam
#
def parsCommandParam(commandsString):
    result = []
    commandsString = commandsString + ' :'

    p = 0
    for i in range(0, len(commandsString)):
        if commandsString[i] == ':':
            p = i

        if commandsString[i] == ' ':
            if commandsString[i +1] == ':':
                subResult = re.sub('\n', '', commandsString[p:i])
                #print (getCommandParam(subResult).command)
                #print (getCommandParam(subResult).param)
                result.append(getCommandParam(subResult))

    return result
# --------------------------------------------------------------------------------------------



