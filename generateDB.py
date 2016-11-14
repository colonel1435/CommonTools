#!usr/bin/env python3.4
# -*- coding: utf-8 -*-
# #  FileName    : generateDB
# #  Author      : Zero
# #  Description : 
# #  Time        : 2016/11/05


import sys
import getopt
import sqlite3
import os
import math
import csv
from  xls2csv import readExcelByRow
from getFileList import createGraphicExcel
from getFileList import createEquipmentExcel
from getFileList import createDepotExcel
from getFileList import getAllDepot

import codecs
import common
from common import Usage
from xls2csv import readExcelByRow

OPTIONS = common.OPTIONS
OPTIONS.json_name = "dumn.json"
OPTIONS.output_dir = None
OPTIONS.turnout_name = u"道岔列表.xlsx"
OPTIONS.all_depot = False
OPTIONS.depot_name = None
OPTIONS.database_postfix = ".db"
OPTIONS.default_dir = "product"
OPTIONS.depotFile = "depotFile.xlsx"
OPTIONS.graphicFile = "graphicFile.xlsx"
OPTIONS.equipmentFile = "equipmentFile.xlsx"

def getDepot(depotFile):
    depotList = readExcelByRow(depotFile)
    print (depotList)
    #depot = [item[0], item[1], item[2] for item in depotList]
    return depotList


def getGraphic(graphicFile):
    depotList = readExcelByRow(graphicFile)
    return depotList
    # return {0: ["山丹_平面布置图.dwg", "0", "PMT"]}


def getEquipment(equipmentFile):
    depotList = readExcelByRow(equipmentFile)
    return depotList
    # return {0: ["DC", "DC"]}


def createDepotTable(conn, depotInfo):
    conn.text_factory = str
    sqlCmd = '''CREATE TABLE Depot
			(id INT PRIMARY KEY NOT NULL,
			 name CHAR(30) NOT NULL,
			 depotDC BLOB NOT NULL,
			 flatFile CHAR(50) NOT NULL);
			 '''
    print(sqlCmd)
    conn.execute(sqlCmd)
    sizeDepot = len(depotInfo)
    for i in range(sizeDepot):
        file = unicode(os.path.join(os.getcwd(), depotInfo[i][2]), "utf-8")
        # file = unicode(os.getcwd() + "\\" + depotInfo[i][2], "utf-8")
        print (file)
        with open(file, 'rb') as f:
            data = f.read()
            insertCmd = '''INSERT INTO Depot (id, name, depotDC, flatFile)
                           VALUES (?, ?, ?, ?)
                        '''
            print(insertCmd)
            conn.execute(insertCmd, (i, depotInfo[i][1], sqlite3.Binary(data), depotInfo[i][3]))
    conn.commit()
'''
    cursor = conn.cursor()
    with open('test.json', 'wb') as out:
        cursor.execute("SELECT depotDC FROM Depot WHERE id = 0")
        ablob = cursor.fetchone()
        out.write(ablob[0])
    cursor.close()
'''

def createGraphicTable(conn, graphicInfo):
    conn.text_factory = str
    sqlCmd = '''CREATE TABLE Graphic
			(gid INT PRIMARY KEY NOT NULL,
			 fileName CHAR(30) NOT NULL,
			 fileData BLOB NOT NULL,
			 type CHAR(30) NOT NULL);'''
    print(sqlCmd)
    conn.execute(sqlCmd)
    sizeGraphic = len(graphicInfo)
    for i in range(sizeGraphic):
        file = unicode(os.path.join(os.getcwd(), graphicInfo[i][2]), "utf-8")
        print (file)
        with open(file, 'rb') as f:
            data = f.read()
            insertCmd = '''
                        INSERT INTO Graphic (gid, fileName, fileData, type)
                        VALUES (?, ?, ? ,?)
                        '''
            print(insertCmd)
            conn.execute(insertCmd, (i, graphicInfo[i][1], sqlite3.Binary(data), graphicInfo[i][3]))
    conn.commit()
'''
    # TODO check output file
    cursor = conn.cursor()
    with open('test.dwg', 'wb') as out:
        cursor.execute("SELECT fileData FROM Graphic WHERE gid = 0")
        ablob = cursor.fetchone()
        out.write(ablob[0])
    cursor.close()
'''
def createEquipmentTable(conn, equipmentInfo):
    sqlCmd = '''CREATE TABLE Equipment
			(eid INT PRIMARY KEY NOT NULL,
			 name CHAR(30) NOT NULL,
			 type CHAR(50) NOT NULL);
	         '''
    print(sqlCmd)
    conn.execute(sqlCmd)
    sizeEquipment = len(equipmentInfo)
    for i in range(sizeEquipment):
        insertCmd = ''' INSERT INTO Equipment (eid, name, type)
                            VALUES ({eid},"{ename}","{etype}")
			    '''.format(eid=i, ename=equipmentInfo[i][1],
                           etype=equipmentInfo[i][2])
        print(insertCmd)
        conn.execute(insertCmd)
    conn.commit()


def createDB(db, depot, graphic, equipment):
    conn = sqlite3.connect(db)
    print("Open db successfully")
    createDepotTable(conn, getDepot(depot))
    createGraphicTable(conn, getGraphic(graphic))
    createEquipmentTable(conn, getEquipment(equipment))
    conn.close()
    print(">>> Create database -> %s successful! <<<" % db)


def main(argv=None):

    def option_handler(p, v):
        if p in ("-d", "--depot"):
            OPTIONS.depot_name = v.decode("gbk")
        elif p == "--all":
            OPTIONS.all_depot = True
        elif p in ("-o", "--output"):
            OPTIONS.output_dir = os.path.join(os.getcwd(), v.decode("gbk"))
        else:
            return False
        return True

    args = common.ParseOptions(argv, __doc__,
                               extra_opts="d:ao:",
                               extra_long_opts=[
                                   "depot=",
                                   "all",
                                   "output="
                               ], extra_opts_handler=option_handler)

    if len(argv) < 1:
        Usage("Please reference help description")
        sys.exit(1)

    if (OPTIONS.depot_name == None) and (OPTIONS.all_depot == False):
        Usage("Please used -d(--depot) or --all")
        sys.exit(1)
    allDepot = None
    if OPTIONS.depot_name:
        allDepot  = [OPTIONS.depot_name]
    else:
        allDepot = getAllDepot(os.getcwd())
    print ("There are " + str(len(allDepot)) + " depots")
    for depot in allDepot:
        productDir = ""
        if OPTIONS.output_dir:
            productDir = os.path.join(OPTIONS.output_dir, depot)
        else:
            productDir = os.path.join(os.getcwd(), OPTIONS.default_dir, depot)
        dbFile = os.path.join(productDir, depot + OPTIONS.database_postfix)

        print (">>> PRODUCT DIR -> " + productDir)
        print (">>> DEPOT -> " + depot)
        print (">>> DATABASE -> " + dbFile)
        if not os.path.exists(productDir):
            os.makedirs(productDir)

        if os.path.exists(dbFile):
            print("%s is existed, delete then rebuild" % dbFile)
            os.remove(dbFile)

        createDepotExcel(OPTIONS.depotFile, productDir, depot)
        createEquipmentExcel(OPTIONS.equipmentFile, productDir, depot)
        createGraphicExcel(OPTIONS.graphicFile, productDir, depot)

        print(">>> Start to create database %s <<<" % dbFile)
        createDB(dbFile,
                 os.path.join(productDir,OPTIONS.depotFile),
                 os.path.join(productDir,OPTIONS.graphicFile),
                 os.path.join(productDir,OPTIONS.equipmentFile))

if __name__ == "__main__":
    main(sys.argv[1:])
