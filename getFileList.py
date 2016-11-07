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

FLAT_POSTFFIX = '平面布置图.dwg'
EQUIPMENT_POSTFFIX = '电缆-设备对应表.xls'
DWG_POSTFFIX = '.dwg'
TYPE_PMT = 'PMT'
TYPE_DLYC = 'DLRK'
TYPE_DLTP = 'DLTP'
TYPE_QJDLTP = 'QJDLTP'
TYPE_JXLL = 'JXLL'
TYPE_QJJXLL = 'QJJXLL'

TYPE_DEVICE_DC = "DC"
TYPE_DEVICE_GDDL = "GDDL"
TYPE_DEVICE_QJGDDL = "QJGDDL"
TYPE_DEVICE_XHJ = "XHJ"
TYPE_DEVICE_QJXHJ = "QJXHJ"
TYPE_DEVICE_BOX = "BOX"
TYPE_DEVICE_QBOX = "QBOX"

TYPE_DEVICE_DC_F = "道岔"
TYPE_DEVICE_GDDL_F = "轨道区段"
TYPE_DEVICE_QJGDDL_F = "区间轨道区段"
TYPE_DEVICE_XHJ_F = "信号机"
TYPE_DEVICE_QJXHJ_F = "区间信号机"
TYPE_DEVICE_BOX_F = "BOX"
TYPE_DEVICE_QBOX_F = "QBOX"

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
    if device == TYPE_DEVICE_GDDL_F:
        return  TYPE_DEVICE_GDDL
    if device == TYPE_DEVICE_QJGDDL_F:
        return TYPE_DEVICE_QJGDDL
    if device == TYPE_DEVICE_XHJ_F:
        return TYPE_DEVICE_XHJ
    if device == TYPE_DEVICE_QJXHJ_F:
        return TYPE_DEVICE_QJXHJ
    return TYPE_DEVICE_BOX
def getDepotList(depot=None):
    fileList = []
    if depot:
        root = os.path.join(os.getcwd(), depot)
    else:
        root = os.getcwd()
    for rootDir, dirNames, fileNames in os.walk(root, topdown=True):
        print ">>> enter dir %s " % rootDir
        for fileName in fileNames:
            if rootDir.find(".git") != -1:
                break
            print fileName
            if fileName.endswith(FLAT_POSTFFIX):
                item = [depot, fileName]
                fileList.append(item)
                break
    print fileList
    return fileList

def getEquipmentList(depot=None):
    depot = depot.encode("utf-8")
    eList = []
    eInnerList = []
    eItemDict = {}
    outputList = []
    out = {}
    eFile = depot + EQUIPMENT_POSTFFIX
    eInnerFile = depot +"区间" + EQUIPMENT_POSTFFIX
    eFileDir = os.path.join(os.getcwd(), depot, "站内", eFile)
    eInnerFileDir = os.path.join(os.getcwd(), depot,  "区间" , eInnerFile)
    eList = readExcelByCol(eFileDir.decode("utf-8"))
    eInnerList = readExcelByCol(eInnerFileDir.decode("utf-8"))
    eList.extend(eInnerList)
    # print eList
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

def getFileList(root):
    fileList = []
    rootDir = os.path.join(os.getcwd(), root)
    sizeCurDir = len(os.getcwd())
    for rootDir, dirNames, fileNames in os.walk(rootDir, topdown=True):
        print ">>> enter dir %s " % rootDir
        for fileName in fileNames:
            if fileName.endswith(DWG_POSTFFIX):
                 dir = os.path.join(unicode(rootDir), fileName)
                 fileDir = dir[sizeCurDir+1:]
                 fileType = getFileType(fileName)
                 item = [fileName, fileDir, fileType]
                 fileList.append(item)

    return fileList

def createDepotExcel(depotFile, depot=None):
    fileDir = os.path.join(os.getcwd(), depotFile)
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
    wSheet.write(0, 2, "flatFile", headerStyle)

    for i in range(3):
        wSheet.col(i).width = 0x0a00 + i * 0x1000
    #fileList = getDepotList()
    #sizeList = len(fileList)
    # print sizeList
    # if sizeList <= 0:
    #     print ">>> No depot file"
    #     return None
    wSheet.write(1, 0, 0)
    wSheet.write(1, 1, depot)
    wSheet.write(1, 2, "_".join([depot,FLAT_POSTFFIX]))

    wBook.save(depotFile)
def createGraphicExcel(file, depot):
    fileDir = os.path.join(os.getcwd(), file)
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
    sizeList = len(fileList)
    print sizeList
    for i in range(1,sizeList):
        wSheet.write(i, 0, i-1)
        wSheet.write(i, 1, fileList[i-1][0])
        wSheet.write(i, 2, fileList[i-1][1])
        wSheet.write(i, 3, fileList[i-1][2])
    wBook.save(file)

def createEquipmentExcel(file, depot):
    fileDir = os.path.join(os.getcwd(), file)
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
        print ">>> There is " + str(sizeList) + "equipments"

    i = 0
    for k, v in fileList.iteritems():
        wSheet.write(i, 0, i)
        eType = getEquipmentType(v)
        wSheet.write(i, 1, k.decode('utf-8'))
        wSheet.write(i, 2, eType.decode('utf-8'))
        i += 1
    wBook.save(file)

def main(argv = None):

    depot = u"山丹"
    depotFile = "depot.xlsx"
    graphicFile = "graphic.xlsx"
    equipmentFile = "equipment.xlsx"
    # createDepotExcel(depotFile, depot)
    # createGraphicExcel(graphicFile, depot)
    createEquipmentExcel(equipmentFile, depot)

if __name__ == "__main__":
    main()