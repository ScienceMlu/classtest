#!/usr/bin/env python
# -*- coding: utf-8 -*-
# -*- coding:utf-8 -*-
import os

'''
遍历主文件下所有的xlsx文件
dom_path 是图形界面entty选出的路径
f_path是路径下所有xlsx文件的绝对路径名，使用前先初始化列表
'''


def findfiles(dom_path, f_path):
    files = os.listdir(dom_path)
    for F in files:
        n_path = dom_path + '/' + F
        if os.path.isfile(n_path):
            if os.path.splitext(n_path)[1] == ".txt":
                f_path.append(n_path)
                print(n_path)
        if os.path.isdir(n_path):
            if F[0] == '.':
                pass
            else:
                findfiles(n_path, f_path)


t = []
path = "E:\java\综合测试"
findfiles(u'%s' % path, t)

'''
for fn in t:
    p, f = os.path.split(fn) # 分离出openpyxl能打开的路径名
    print("dir is:" + p)
    print("file is:" + f)
'''