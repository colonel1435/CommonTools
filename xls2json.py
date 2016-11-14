#!usr/bin/env python3.4
# -*- coding: utf-8 -*-
# #  FileName    : xls2json
# #  Author      : Administrator
# #  Description : 
# #  Time        : 2016/11/10

import sys
import os
import json
import codecs
import common
from common import Usage
from xls2csv import readExcelByRow

DUMN_OPTIONS = common.OPTIONS
DUMN_OPTIONS.json_name = "dumn.json"
DUMN_OPTIONS.turnout_name = u"道岔列表.xlsx"


def getTurnoutFile(depot):
    rootDir = ""
    turnoutFile = ""
    if depot == None:
        return turnoutFile;
    else:
        rootDir = os.path.join(os.getcwd(), DUMN_OPTIONS.depot_name)
    for root, dirNames, fileNames in os.walk(rootDir, topdown=True):
        print (">>> Enter %s" % root)
        for fileName in fileNames:
            if fileName == DUMN_OPTIONS.turnout_name:
                turnoutFile = os.path.join(unicode(root), fileName)
                print (turnoutFile)
                return turnoutFile
    return turnoutFile

def getJsonItems():
    jsonList = []
    itemList = []
    excleDir = getTurnoutFile(DUMN_OPTIONS.depot_name)
    if len(excleDir) == 0:
        print ("None excle")
        return jsonList
    itemList = readExcelByRow(excleDir, 1)
    for i in range(len(itemList)):
        jsonList.append(itemList[i][1].rstrip(";").split(";"))
        jsonList[i].insert(0, str(itemList[i][0]))
    return jsonList

def createJson(depot, outputDir=None):
    jsonList = []
    jsonDict = {}
    output_dict = {}
    DUMN_OPTIONS.depot_name = depot
    jsonDir = os.path.join(outputDir, DUMN_OPTIONS.json_name)
    if os.path.exists(jsonDir):
        print (">>> Del old json file then recreate it ...")
        os.remove(jsonDir)
    jsonItems = getJsonItems()
    lenItems = len(jsonItems)
    if lenItems == 0:
        return "";
    jsonDict = {item[0]:item[1:] for item in jsonItems}
    jsonList.append(jsonDict)
    output_dict[depot] = jsonDict
    print (output_dict)
    with codecs.open(jsonDir, 'w', "utf-8") as f:
        f.write(json.dumps(output_dict, ensure_ascii=False))
    return jsonDir

def main(argv=None):
    DUMN_OPTIONS.depot_name = u"高台南".encode("utf-8")
    print (DUMN_OPTIONS.depot_name)
    DUMN_OPTIONS.output_dir = os.path.join(DUMN_OPTIONS.output_dir, DUMN_OPTIONS.depot_name)
    print (DUMN_OPTIONS.output_dir)
    createJson()
if __name__ == "__main__":
    main(sys.argv[1:])