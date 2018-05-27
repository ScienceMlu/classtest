#! /usr/bin/python
# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Border, Side, Font


def set_border(cell_range):
    rows = cell_range
    side = Side(border_style='medium', color="FF000000")
    side2 = Side(border_style='thin', color='FF000000')
    # font = Font(name='', size=12,border=True)
    rows = list(rows)
    max_y = len(rows) - 1
    for pos_y, cells in enumerate(rows):
        max_x = len(cells) - 1  # pos_y 行 pos_x 列
        for pos_x, cell in enumerate(cells):
            border = Border(
                left=cell.border.left,
                right=cell.border.right,
                top=cell.border.top,
                bottom=cell.border.bottom
            )
            if pos_x == 0:
                border.left = side
            if pos_x == max_x:
                border.right = side
            if pos_y == 0:
                border.top = side
            if pos_y == max_y:
                border.bottom = side
            # 内部格式
            if pos_x != 0 and pos_x != max_x:
                border.right = side
                border.left = side
            if pos_y != 0 and pos_y != max_y:
                border.top = side2
                border.bottom = side2
            # if pos_x == 0 or pos_x == max_x or pos_y == 0 or pos_y == max_y:
            cell.border = border


wb = openpyxl.load_workbook('1.xlsx')
ws = wb['Sheet2']
cellrange = ws['C5:I20']
set_border(cellrange)
wb.save('1.xlsx')