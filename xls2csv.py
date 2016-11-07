#!usr/bin/env python3.4
# -*- coding: utf-8 -*-
# #  FileName    : xls2csv
# #  Author      : Zero
# #  Description : 
# #  Time        : 2016/11/05

import xlrd
import xlwt
import sys
import types
from datetime import date, datetime

def cell2csv():
    return None

def readExcel(excel):
    xlrd.Book.encoding='utf-8'
    sheetInfo = []
    workBook = xlrd.open_workbook(excel)
    sheet1 = workBook.sheet_by_index(0)
    for row in range(1, sheet1.nrows):
        rows = sheet1.row_values(row)
        def _2str(cell):
            # return cell.encode('gbk')
            if type(u'') == type(cell) or type('') == type(cell):
                return cell.encode("utf-8")
            else:
                return str(cell)
        sheetInfo.append([_2str(cell) for cell in rows])
    return sheetInfo