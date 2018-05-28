#! /usr/bin/python
# -*- coding: utf-8 -*-

import os
import shutil
import string


def find_dom_name(dom_path):
    dom_name = dom_path.split('\\')[-1]
    return dom_name


def find_sheet_name(dom_path):
    sheetname = []
    files = os.listdir(dom_path)
    for f in files:
        n_path = dom_path + '\\' + f
        if os.path.isdir(n_path):
            sheetname.append(f)
    return sheetname


def move_file(sourceDir, targetDir):
    shutil.copyfile(sourceDir, targetDir)


def findfiles(dom_path, tc_real_path, tc_name, target):
    files = os.listdir(dom_path)
    for F in files:
        n_path = dom_path + '\\' + F
        if os.path.isfile(n_path):
            if os.path.splitext(n_path)[-1] == '.txt':
                sStr = n_path.replace('\\', '_').replace(':', '')
                tc_name.append(sStr)
                final_target = '%s\\%s' % (target, sStr)
                move_file(n_path, final_target)
                tc_real_path.append(n_path)
        if os.path.isdir(n_path):
            if F[0] == '.':
                pass
            else:
                findfiles(n_path, tc_real_path, tc_name, target)


t = []
tcname = []
path = r'E:\java\综合测试'
target_path = r'E:\java\结果'

findfiles(path, t, tcname, target_path)

# findfiles(path, t, cf, target_path)
str1 = '["Sheet1：[X36]","Sheet3：[R52]"]'
print(str1.strip('[' + ']'))
print(str1.replace('["', '').replace('"]', '').replace('"', ''))  # 去除强制转化字符串所带来的垃圾标点符号
