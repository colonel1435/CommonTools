#!usr/bin/env python3.4
# -*- coding: utf-8 -*-
# #  FileName    : generateDB
# #  Author      : Zero
# #  Description :
# #  Time        : 2016/11/05

import os
import sys
import xlwt
import xlrd
from xlutils.copy import copy
from  xls2csv import readExcelByCol
from xls2json import createJson

FLAT_POSTFFIX = '_平面布置图.dwg'
EQUIPMENT_POSTFFIX = '电缆-设备对应表.xls'
DWG_POSTFFIX = '.dwg'
TYPE_PMT = 'PMT'
TYPE_DLYC = 'DLRK'
TYPE_DLTP = 'DLTP'
TYPE_QJDLTP = 'QDLTP'
TYPE_JXLL = 'JXLL'
TYPE_QJJXLL = 'QJXLL'

TYPE_DEVICE_DC = "DC"
TYPE_DEVICE_GDDL = "GDDL"
TYPE_DEVICE_QJGDDL = "QJGDDL"
TYPE_DEVICE_XHJ = "XHJ"
TYPE_DEVICE_QJXHJ = "QJXHJ"
TYPE_DEVICE_BOX = "BOX"
TYPE_DEVICE_QBOX = "QBOX"
TYPE_DEVICE_LK = "LK"

TYPE_DEVICE_DC_F = "道岔"
TYPE_DEVICE_GDDL_F = "轨道区段"
TYPE_DEVICE_QJGDDL_F = "区间轨道区段"
TYPE_DEVICE_XHJ_F = "信号机"
TYPE_DEVICE_QJXHJ_F = "区间信号机"
TYPE_DEVICE_DCARXHJ_F = "调车信号机"
TYPE_DEVICE_LCARXHJ_F = "列车信号机"
TYPE_DEVICE_BOX_F = "BOX"
TYPE_DEVICE_QBOX_F = "QBOX"
TYPE_DEVICE_LK_F  = "列控"

black_list = ["product", ".git"]

def getFileType(file):
    if file.find(TYPE_DLYC) != -1:
        return TYPE_DLYC
    if file.find(TYPE_QJDLTP) != -1:
        return TYPE_QJDLTP
    if file.find(TYPE_QJJXLL) != -1:
        return TYPE_QJJXLL
    if file.find(TYPE_DLTP) != -1:
        return TYPE_DLTP
    if file.find(TYPE_JXLL) != -1:
        return TYPE_JXLL
    else:
        return TYPE_PMT

def getEquipmentType(device):
    if device == TYPE_DEVICE_DC_F:
        return TYPE_DEVICE_DC
    if device == TYPE_DEVICE_QJGDDL_F:
        return TYPE_DEVICE_QJGDDL
    if device == TYPE_DEVICE_GDDL_F:
        return  TYPE_DEVICE_GDDL
    if device == TYPE_DEVICE_QJXHJ_F:
        return TYPE_DEVICE_QJXHJ
    if device == TYPE_DEVICE_XHJ_F:
        return TYPE_DEVICE_XHJ
    if device == TYPE_DEVICE_DCARXHJ_F:
        return TYPE_DEVICE_XHJ
    if device == TYPE_DEVICE_LCARXHJ_F:
        return TYPE_DEVICE_XHJ
    if device == TYPE_DEVICE_QBOX_F:
        return TYPE_DEVICE_QBOX
    if device == TYPE_DEVICE_BOX_F:
        return TYPE_DEVICE_BOX
    if device == TYPE_DEVICE_LK_F:
        return TYPE_DEVICE_LK
    return TYPE_DEVICE_BOX

def getAllDepot(root):
    depotList = []
    for rootDir, dirNames, fileNames in os.walk(root.decode("utf-8"), topdown=True):
        for item in dirNames:
            if item not in black_list:
                depotList.append(item)
        if depotList:
            return depotList
    return depotList

def getAllETable(root):
    eList = []
    for rootDir, dirNames, fileNames in os.walk(root, topdown=True):
        for fileName in fileNames:
            if fileName.endswith(EQUIPMENT_POSTFFIX.decode("utf-8")):
                eFile = os.path.join(rootDir, fileName)
                eList.append(eFile)
    return eList
def getDepotDC(depot, outputDir):
    if depot == None:
        return None;
    dcDir = createJson(depot, outputDir)
    return dcDir

def getDepotList(depot=None, outputDir=None):
    fileList = []
    root = None
    depotList = []
    if depot:
        root = os.path.join(os.getcwd(), depot)
    else:
        return fileList
    lenCurDir = len(os.getcwd())
    for rootDir, dirNames, fileNames in os.walk(root, topdown=True):
        print "> Enter dir %s " % rootDir
        for fileName in fileNames:
            flat = (depot + FLAT_POSTFFIX.decode("utf-8"))
            if fileName == flat:
            # if fileName.endswith(FLAT_POSTFFIX):
                dcDir = getDepotDC(depot, outputDir)
                item = [depot, dcDir[lenCurDir:], fileName]
                fileList.append(item)
                return fileList
    # print fileList
    return fileList


def getEquipmentList(depot=None, outputDir=None):
    eList = []
    eInnerList = []
    eItemDict = {}
    outputList = []
    out = {}

    eTableList = getAllETable(depot)
    # eFile = depot + EQUIPMENT_POSTFFIX.decode("utf-8")
    # eInnerFile = depot +u"区间" + EQUIPMENT_POSTFFIX.decode("utf-8")
    # e2216File = u"2216区间" + EQUIPMENT_POSTFFIX.decode("utf-8")
    # e2236File = u"2236区间" + EQUIPMENT_POSTFFIX.decode("utf-8")
    #
    #
    # eFileDir = os.path.join(os.getcwd(), depot, depot, eFile)
    # eInnerFileDir = os.path.join(os.getcwd(), depot,  depot+u"区间" , eInnerFile)
    # e2216Dir = os.path.join(os.getcwd(), depot, u"中继站2216", e2216File)
    # e2236Dir = os.path.join(os.getcwd(), depot, u"中继站2236", e2236File)
    for e in eTableList:
        if isinstance(e, unicode):
            print ("#### UNICODE #####")
        eList.extend(readExcelByCol(e))
    # eList = readExcelByCol(eFileDir)
    # eList.extend(readExcelByCol(eInnerFileDir))
    # eList.extend(readExcelByCol(e2216Dir))
    # eList.extend(readExcelByCol(e2236Dir))
    print eList
    for item in eList:
        for i in range(1, len(item)) :
            if item[i].strip():
                if eItemDict.has_key(item[0]):
                    eItemDict[item[0]].append(item[i])
                else:
                    eItemDict[item[0]] = [item[i]]
    for k, v in eItemDict.iteritems():
        for item in v:
            outputList.append({e:k for e in item.split(';') if e.strip()})

    for i in range(len(outputList)):
        # print outputList[i]
        if i == 0:
            out = outputList[0]
        else:
            out = dict(out, **outputList[i])
    return out

def getFileList(depot=None, outputDir=None):
    fileList = []
    root = None
    if depot:
        root = os.path.join(os.getcwd(), depot)
    else:
        return fileList
    sizeCurDir = len(os.getcwd())
    for rootDir, dirNames, fileNames in os.walk(root, topdown=True):
        print ">>> Enter dir %s " % rootDir
        for fileName in fileNames:
            if fileName.endswith(DWG_POSTFFIX):
                 dir = os.path.join(unicode(rootDir), fileName)
                 fileDir = dir[sizeCurDir+1:]
                 fileType = getFileType(fileName)
                 item = [fileName, fileDir, fileType]
                 fileList.append(item)

    return fileList

def createDepotExcel(depotFile, outputDir=None, depot=None):
    print(">>> Start to create depot excel")
    fileDir = os.path.join(outputDir, depotFile)
    if os.path.exists(fileDir):
        print ">>> %s has already existed, delete then recreate..." % fileDir
        os.remove(fileDir)

    xlrd.Book.encoding = 'utf-8'
    wBook = xlwt.Workbook()
    wSheet = wBook.add_sheet('sheet1', cell_overwrite_ok=True)
    styleBoldRed = xlwt.easyxf('font: color-index black, bold on')
    headerStyle = styleBoldRed
    wSheet.write(0, 0, "id", headerStyle)
    wSheet.write(0, 1, "name", headerStyle)
    wSheet.write(0, 2, "depotDC", headerStyle)
    wSheet.write(0, 3, "flatFile", headerStyle)

    for i in range(4):
        wSheet.col(i).width = 0x0a00 + i * 0x1000

    depotList = getDepotList(depot, outputDir)
    sizeList = 0
    if depotList == None:
        print ">>> No depot file"
        return None
    else:
        sizeList = len(depotList)
        print (">>> This is " + depot + " depot")

    for i in range(sizeList):
        wSheet.write(i, 0, 0)
        wSheet.write(i, 1, depotList[i][0])
        wSheet.write(i, 2, depotList[i][1])
        wSheet.write(i, 3, depotList[i][2])
    wBook.save(fileDir)
def createGraphicExcel(file, outputDir = None, depot=None):
    print(">>> Start to create graphic excel")
    fileDir = os.path.join(outputDir, file)
    if os.path.exists(fileDir):
        print ">>> %s has already existed, delete then recreate..." % fileDir
        os.remove(fileDir)
    xlrd.Book.encoding = 'utf-8'
    wBook = xlwt.Workbook()
    wSheet = wBook.add_sheet('sheet1', cell_overwrite_ok=True)
    styleBoldRed = xlwt.easyxf('font: color-index black, bold on')
    headerStyle = styleBoldRed
    wSheet.write(0, 0, "gid", headerStyle)
    wSheet.write(0, 1, "fileName", headerStyle)
    wSheet.write(0, 2, "fileData", headerStyle)
    wSheet.write(0, 3, "type", headerStyle)

    for i in range(3):
        wSheet.col(i).width = 0x0a00 + i * 0x2000

    fileList = getFileList(depot)
    sizeList = 0
    if fileList == None:
        print "There is nothing"
        return None
    else:
        sizeList = len(fileList)
        print (">>> There are " + str(sizeList) + " dwgs")

    for i in range(1,sizeList):
        wSheet.write(i, 0, i-1)
        wSheet.write(i, 1, fileList[i-1][0])
        wSheet.write(i, 2, fileList[i-1][1])
        wSheet.write(i, 3, fileList[i-1][2])
    wBook.save(fileDir)

def createEquipmentExcel(file, outputDir = None, depot = None):
    print(">>> Start to create equipment excel")
    fileDir = os.path.join(outputDir, file)
    if os.path.exists(fileDir):
        print ">>> %s has already existed, delete then recreate..." % fileDir
        os.remove(fileDir)
    xlrd.Book.encoding = 'utf-8'
    wBook = xlwt.Workbook()
    wSheet = wBook.add_sheet('sheet1', cell_overwrite_ok=True)
    styleBoldRed = xlwt.easyxf('font: color-index black, bold on')
    headerStyle = styleBoldRed
    wSheet.write(0, 0, "eid", headerStyle)
    wSheet.write(0, 1, "name", headerStyle)
    wSheet.write(0, 2, "type", headerStyle)

    for i in range(3):
        wSheet.col(i).width = 0x0a00 + i*0x1000

    fileList = getEquipmentList(depot)
    if fileList == None:
        print "There is nothing"
        return None
    else:
        sizeList = len(fileList)
        print ">>> There are " + str(sizeList) + " equipments"

    i = 0
    for k, v in fileList.iteritems():
        wSheet.write(i, 0, i)
        eType = getEquipmentType(v)
        wSheet.write(i, 1, k.decode('utf-8'))
        wSheet.write(i, 2, eType.decode('utf-8'))
        i += 1
    wBook.save(fileDir)

def main(argv = None):
    depot = "高台南"
    if depot != None:
        productDir = os.path.join(os.getcwd(), "product", depot)
    else:
        productDir = os.path.join(os.getcwd(), "product")
    print (">>> PRODUCT DIR -> " + productDir)
    print (">>> DEPOT -> " + depot)
    if isinstance(productDir, unicode):
        print "### UNICODE ###"
    if not os.path.exists(productDir.decode("utf-8")):
        os.makedirs(productDir.decode("utf-8"))

    depotFile = "depotFile.xlsx"
    graphicFile = "graphicFile.xlsx"
    equipmentFile = "equipmentFile.xlsx"

    createDepotExcel(depotFile, depot, productDir)
    createEquipmentExcel(equipmentFile, depot, productDir)
    createGraphicExcel(graphicFile, depot, productDir)

if __name__ == "__main__":
    main()