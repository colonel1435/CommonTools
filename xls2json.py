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

OPTIONS = common.OPTIONS
OPTIONS.json_name = "dumn.json"
OPTIONS.output_dir = os.getcwd()
OPTIONS.turnout_name = u"道岔列表.xlsx"
OPTIONS.append_depot = False
OPTIONS.depot_name = None


def getTurnoutFile(depot = None):
    rootDir = ""
    turnoutFile = ""
    if depot == None:
        rootDir = os.getcwd();
    else:
        rootDir = os.path.join(os.getcwd(), OPTIONS.depot_name)
    for root, dirNames, fileNames in os.walk(rootDir, topdown=True):
        print (">>> Enter %s" % root)
        for fileName in fileNames:
            if fileName == OPTIONS.turnout_name:
                turnoutFile = os.path.join(unicode(root), fileName)
                print (turnoutFile)
                return turnoutFile
    return turnoutFile

def getJsonItems(jsonDir):
    jsonList = []
    itemList = []
    excleDir = getTurnoutFile(OPTIONS.depot_name)
    if len(excleDir) == 0:
        print ("None excle")
        return jsonList
    itemList = readExcelByRow(excleDir, 1)
    for i in range(len(itemList)):
        jsonList.append(itemList[i][1].rstrip(";").split(";"))
        jsonList[i].insert(0, str(itemList[i][0]))
    return jsonList

def createJson(depot=None, outputDir=None):
    if depot != None and outputDir != None:
        OPTIONS.depot_name = depot
        OPTIONS.output_dir = outputDir
    jsonList = []
    jsonDict = {}
    output_dict = {}
    if (not str(os.path.exists(OPTIONS.output_dir)).decode("utf-8")):
        print (">>> No depot data")
        return;
    jsonDir = os.path.join(OPTIONS.output_dir, OPTIONS.json_name)
    if os.path.exists(jsonDir):
        print (">>> Del old json file then recreate it ...")
        os.remove(jsonDir)
    jsonItems = getJsonItems(jsonDir)
    lenItems = len(jsonItems)
    if lenItems == 0:
        return;
    jsonDict = {item[0]:item[1:] for item in jsonItems}
    jsonList.append(jsonDict)
    output_dict[OPTIONS.depot_name] = jsonDict
    print (output_dict)
    with codecs.open(jsonDir, 'w', "utf-8") as f:
        f.write(json.dumps(output_dict, ensure_ascii=False))
    return jsonDir

def main(argv=None):
    def option_handler(p, v):
        if p in ("-d", "--depot"):
            OPTIONS.depot_name = v.decode("gbk")
        elif p in ("-a", "--append"):
            OPTIONS.append = True
        elif p in ("-o", "--output"):
            OPTIONS.output_dir = v.decode("gbk")
        else:
            return False
        return True

    # args = common.ParseOptions(argv, __doc__,
    #                            extra_opts="d:ao:",
    #                            extra_long_opts=[
    #                                "depot=",
    #                                "append=",
    #                                "output="
    #                            ], extra_opts_handler=option_handler)
    #
    # if OPTIONS.append_depot == True:
    #     Usage("No cmd")
    #     pass
    # elif OPTIONS.depot_name != None:
    #     print (OPTIONS.depot_name)
    #     OPTIONS.output_dir = os.path.join(OPTIONS.output_dir, OPTIONS.depot_name)
    #     print (OPTIONS.output_dir)
    #     createJson()
    # else:
    #     Usage("Please enter depot name with -d or --depot")
    #     sys.exit(1)
    OPTIONS.depot_name = u"高台南".encode("utf-8")
    print (OPTIONS.depot_name)
    OPTIONS.output_dir = os.path.join(OPTIONS.output_dir, OPTIONS.depot_name)
    print (OPTIONS.output_dir)
    createJson()
if __name__ == "__main__":
    main(sys.argv[1:])