#! /usr/bin/python
# -*- coding: utf-8 -*-
import openpyxl
import re


def check_timetable(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet_name = wb.sheetnames
    for j in sheet_name:
        if j == '時間＿承認':
            sheet1 = wb[j]
            '''此处写死'''
            if sheet1['C9'].value is None:
                print()
                data_error_txt = '请填写检查时间（使用Ctrl+；）'
            if not re.match(u'^\u672a\u67e5\u8be2\u5230\u7ed3\u679c', sheet1['C13'].value):
                name_error_txt = '请填写检查人姓名'
            if sheet1['C15'].value is None:
                version_error_txt = '请输入版本号'
            if sheet1['C17'].value is None:
                pcname_error_txt = '请输入测试pc名'
            if sheet1['C19'].value is None:
                osAndlang_error_txt = '请输入pc的os名和言语环境'
            if not re.match(r'WLAN'|'Bluetooth'|'USB',sheet1['C21']):
                IF_error_txt = '请输入连接方式'
            if sheet1['C23'].value is None:
                Model_error_txt = '请输入评价设备名'
            if sheet1['C26'].value is None and not re.match('^\d', sheet1['C26'].value):
                time_error_txt = '请输入评价时间'
