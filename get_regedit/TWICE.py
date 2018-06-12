#! /usr/bin/python
# -*- coding: utf-8 -*-
import winreg
class regedit:
    def __init__(self, path, mainfile):
        self.path = path
        self.mainfile = mainfile
    def print_value(self):
        key = winreg.OpenKey(self.mainfile, self.path)
        countkey = winreg.QueryInfoKey(key)[1]
        for i in range(int(countkey)):
            name = winreg.EnumValue(key, i)  # 获取数据，【0】：名称 【1】类型 【2】：数值
            value = winreg.QueryValueEx(key, name[0])
            print('名称：'+str(name[0])+','+'数据：'+ str(value[0]))

# path = r"SYSTEM\CurrentControlSet\Control\Print\Printers\Microsoft Print to PDF\PrinterDriverData"
#  mainfile = winreg.HKEY_LOCAL_MACHINE

regedit1 =regedit(path=r"SYSTEM\CurrentControlSet\Control\Print\Printers\Microsoft Print to PDF\PrinterDriverData", mainfile=winreg.HKEY_LOCAL_MACHINE)
regedit1.print_value()
