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
from  xls2csv import readExcel
from getFileList import createGraphicExcel

class Usage(Exception):
    def __init__(self, msg):
        self.msg = msg


def getDepot(depotFile):
    depotList = readExcel(depotFile)
    #depot = [item[0], item[1], item[2] for item in depotList]
    return depotList


def getGraphic(graphicFile):
    depotList = readExcel(graphicFile)
    return depotList
    # return {0: ["山丹_平面布置图.dwg", "0", "PMT"]}


def getEquipment(equipmentFile):
    depotList = readExcel(equipmentFile)
    return depotList
    # return {0: ["DC", "DC"]}


def createDepotTable(conn, depotInfo):
    sqlCmd = '''CREATE TABLE Depot
			(id INT PRIMARY KEY NOT NULL,
			 name CHAR(30) NOT NULL,
			 flatName CHAR(50) NOT NULL);
			 '''
    print(sqlCmd)
    conn.execute(sqlCmd)
    sizeDepot = len(depotInfo)
    for i in range(sizeDepot):
        insertCmd = '''INSERT INTO Depot (id, name, flatName)
                               VALUES (%d, "%s", "%s")
			    ''' % (i, depotInfo[i][1], depotInfo[i][2])
        print(insertCmd)
        conn.execute(insertCmd)
    conn.commit()


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
            # insertCmd = '''INSERT INTO Graphic (gid, fileName, fileData, type)
            #         VALUES ({gid},"{fileName}",{fileData},"{fileType}")
            #         '''.format(gid=i, fileName=graphicInfo[i][1],
            #                    fileData=sqlite3.Binary(data),
            #                    fileType=graphicInfo[i][3])
            insertCmd = '''
                        INSERT INTO Graphic (gid, fileName, fileData, type)
                        VALUES (?, ?, ? ,?)
                        '''
            print(insertCmd)
            # conn.execute(insertCmd)
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


def createDB(db):
    depotFile = "depotFile.xlsx"
    graphicFile = "graphicFile.xlsx"
    equipmentFile = "equipmentFile.xlsx"
    depot = u"山丹"
    createGraphicExcel(graphicFile, depot)
    dbFile = os.getcwd() + "\\" + db
    if os.path.exists(dbFile):
        print("%s is existed, delete then rebuild" % db)
        os.remove(dbFile)
    conn = sqlite3.connect(db)
    print("Open db successfully")
    createDepotTable(conn, getDepot(depotFile))
    createGraphicTable(conn, getGraphic(graphicFile))
    createEquipmentTable(conn, getEquipment(equipmentFile))
    conn.close()
    print(">>> Create database -> %s successful! <<<" % db)


def main(argv=None):
    dbFile = "sunland.db"
    print(">>> Start to create database %s <<<" % dbFile)
    createDB(dbFile)


if __name__ == "__main__":
    #        sys.exit(main())
    # reload(sys)
    # sys.setdefaultencoding('utf8')

    main()
