#! /usr/bin/python
# -*- coding: utf-8 -*-

import findpath
import checkTCwrite
import os
import shutil
path = input('请输入文件夹路径：')

findpath.get_file_name(path)


# 创建结果文件夹
try:
    os.mkdir('end')
except FileExistsError:
    print('文件夹已存在')
    pass


for i in range(len(findpath.excelfile)):
    t = checkTCwrite.check_tc(findpath.excelfile[i])

    # str1 = "Result0.xlsx"
    str1 = "%s.Result.xlsx" % (findpath.excelfile[i])

    checkTCwrite.write_result(t[0], t[1], str1)
    shutil.move(str1, 'end')
