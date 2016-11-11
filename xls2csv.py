#!usr/bin/env python3.4
# -*- coding: utf-8 -*-
# #  FileName    : xls2csv
# #  Author      : Zero
# #  Description : 
# #  Time        : 2016/11/05

import xlrd
import xlwt
import sys
import os
import types
from datetime import date, datetime

def cell2csv():
    return None

def readExcelByRow(excel, row=1):
    xlrd.Book.encoding='utf-8'
    sheetInfo = []
    workBook = xlrd.open_workbook(excel)
    sheet1 = workBook.sheet_by_index(0)
    for row in range(row, sheet1.nrows):
        rows = sheet1.row_values(row)
        def _2str(cell):
            if type(u'') == type(cell) or type('') == type(cell):
                return cell.encode("utf-8")
            else:
                return str(int(cell))
        sheetInfo.append([_2str(cell) for cell in rows])
    return sheetInfo

def readExcelByCol(excel, col=1):
    if os.path.exists(excel):
        print ">>> %s is existed" % excel
    else:
        print ">>> %s is nothing" % excel
    xlrd.Book.encoding = "utf-8"
    sheetInfo = []
    workBook = xlrd.open_workbook(excel)
    sheet1 = workBook.sheet_by_index(0)
    for col in range(col, sheet1.ncols):
        cols = sheet1.col_values(col)
        def _2str(cell):
            if type(u'') == type(cell) or type('') == type(cell):
                return cell.encode("utf-8")
            else:
                return str(cell)
        sheetInfo.append([_2str(cell) for cell in cols])
    return sheetInfo