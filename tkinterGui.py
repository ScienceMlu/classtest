#!/usr/bin/env python
# -*- coding: utf-8 -*-
import shutil
import tkinter.filedialog as filedialog
import os

from past.builtins import raw_input
import checkTCwrite
from tkinter import *


def callback():
    entry.delete(0, END)  # 清空entry里面的内容
    listbox_filename.delete(0, END)
    # 调用filedialog模块的askdirectory()函数去打开文件夹
    global file_path
    file_path = filedialog.askdirectory()
    if file_path:
        entry.insert(0, file_path)  # 将选择好的路径加入到entry里面
    # print(file_path)
    # copy_excel_file(file_path)
    getdir(file_path)


def getdir(file_document_path=os.getcwd()):
    listDemo = []
    """
    用于获取目录下的文件列表
    """
    cf = os.listdir(file_document_path)

    try:
        os.mkdir('end')
    except FileExistsError:
        print('文件夹已存在')
        doc_exist_code = True

    for i in cf:
        # os.path.splitext():分离文件名与扩展名
        if os.path.splitext(i)[1] == '.xlsx':
            listbox_filename.insert(END, i)
            listDemo.append(i)
            print(listDemo)
    for j in range(len(listDemo)):
        t = checkTCwrite.check_tc(listDemo[j])
        # str1 = "Result0.xlsx"
        str1 = "%s.Result.xlsx" % listDemo[j]
        '''
        需要插入异常处理，防止结果文件夹内容未删除导致报错，并且赋值给特定变量，指向gui警告弹框
        '''
        checkTCwrite.write_result(t[0], t[1], str1)
        shutil.move(str1, 'end')


def copy_excel_file(file_real_path):
    work_path = os.getcwd()
    all_list = os.listdir(file_real_path)
    print(work_path, file_real_path)
    for i in all_list:
        aa, bb = i.split('.')

        '''
        if 'xlsx' in bb.lower():
            oldname = file_path+'//'+aa+'.'+bb
            newname = work_path+'//'+aa+'.'+bb
            shutil.copyfile(oldname,newname)
        '''


if __name__ == "__main__":
    root = Tk()
    root.title("TC审查工具")
    root.geometry("400x400")
    root.rowconfigure(1, weight=1)
    root.rowconfigure(2, weight=8)

    entry = Entry(root, width=60)
    entry.grid(sticky=W + N, row=0, column=0, columnspan=4, padx=5, pady=5)
    button = Button(root, text="选择文件夹并开始", command=callback)
    # button2 = Button(root, text="开始").grid(row=1,column=2)  # command=cosoleA.sh_check(file_path))
    button.grid(sticky=W + N, row=1, column=0, padx=5, pady=5)
    # 创建listbox用来显示所有文件名
    listbox_filename = Listbox(root, width=60)
    listbox_filename.grid(row=2, column=0, columnspan=4, rowspan=4,
                          padx=5, pady=5, sticky=W + E + S + N)
    root.mainloop()
    raw_input()

