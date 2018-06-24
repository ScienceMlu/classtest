#! /usr/bin/python
# -*- coding: utf-8 -*-
from classtest import exceldata_move

t1 = exceldata_move.BaseUse(filename='1.xlsx', sheetname='Sheet1')
#print(t1.keyRow_catch(keyword='tools')[1])
t2 = t1.cols_catch_row(column='B',keyword_us='tools')
print(t2)