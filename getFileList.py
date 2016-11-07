#!usr/bin/env python3.4
# -*- coding: utf-8 -*-
# #  FileName    : generateDB
# #  Author      : Zero
# #  Description :
# #  Time        : 2016/11/05

import os
import xlwt
import xlrd
from xlutils.copy import copy

FLAT_POSTFFIX = u'平面布置图.dwg'
DWG_POSTFFIX = '.dwg'

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
                 fileType = "PMT"
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
    # for i in range(sizeList):
    #     ws.write(i, j, fileList)

    wBook.save(file)


def main(argv = None):
    depot = u"山丹"
    depotFile = "depot.xlsx"
    graphicFile = "graphic.xlsx"
    equipmentFile = "equipment.xlsx"
    createDepotExcel(depotFile, depot)
    # createGraphicExcel(graphicFile, depot)

if __name__ == "__main__":
    main()