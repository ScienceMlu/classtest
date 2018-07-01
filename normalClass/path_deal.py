#! /usr/bin/python
# -*- coding: utf-8 -*-
import os
import shutil
from shutil import move


class FindFile:
    def __init__(self, absolute_path, extension_word):
        self.absolute_document_path = absolute_path
        self.extension_txt = extension_word
        self.allfile_list = []

    def separate_extensiontxt_mode1(self):  # 指定目录下,寻找对应后缀名，并分离文件和扩展名
        files = os.listdir(self.absolute_document_path)
        filename_list = []
        for file_value in files:
            if file_value.split('.', -1) == self.extension_txt:
                filename_list.append(file_value.split('.', 0))
        return filename_list

    def copyAndmove_file(self, oldfile_absoulte_name, new_absolute_name):  # 复制文件并且重命名
        shutil.copyfile(oldfile_absoulte_name, new_absolute_name)

    def reformat_path(self, path, old_keymark, new_keymark): # 文件重命名
        new_path = path.replace(old_keymark, new_keymark)
        return new_path

    def find_allfile_inDir(self, document_path):  # 递归遍历寻找主目录下所有文件，列表输出
        files = os.listdir(document_path)
        for file_value in files:
            n_path = self.absolute_document_path + '\\' + file_value
            if os.path.isfile(n_path):
                self.allfile_list.append(n_path)
            if os.path.isdir(n_path):
                self.find_allfile_inDir(n_path)
            else:
                raise FileNotFoundError

    def create_resultDir(self, work_path, dirname):  # 创建结果文件夹，防止重复创建
        os.chdir(work_path)
        file = os.listdir(work_path)
        for file_value in file:
            if file_value == dirname:
                os.remove(dirname)
                os.mkdir(dirname)
            else:
                os.mkdir(dirname)



