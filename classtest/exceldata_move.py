#! /usr/bin/python
# -*- coding: utf-8 -*-
from openpyxl import load_workbook


class BaseUse(object):
    def __init__(self, filename, keyword):
        self.filename = filename
        self.keyword = keyword
        self.wb = load_workbook(filename)
        self.ws = self.wb['outlook']

    def cell_format(self):
        for row in self.ws.max_row:
            self.ws.row_dimensions[row].height = 83

    def keyword_row(self):
        for u_row in self.ws.iter_rows():
            for cell in u_row:
                if cell.value == self.keyword:
                    return cell.row

    def old_data_catch(self):
        data_list = []
        Nomax = self.ws.max_row - self.keyword_row()
        for c_row in self.ws.iter_rows():
            for cell in c_row:
                if cell.value == self.keyword:
                    for No in range(Nomax):
                        data_list.append(self.ws['%s%d' % (cell.column, cell.row + No + 1)].value)
        return data_list


class EDataMix(BaseUse):
    def __init__(self, filename, keyword=None, mix_format=None):
        BaseUse.__init__(self, filename, keyword)
        self.mix_format = mix_format

    def mix_2(self, data_list_base, data_list_use):
        mix_data = []
        for data_key, data_value in enumerate(data_list_base):
            mix_data.append(self.mix_format % (data_value, data_list_use[data_key]))
        print(mix_data)
        return mix_data

    def mix_3(self, data_list_base, data_list_use1, data_list_use2):
        mix_data = []
        for data_key, data_value in enumerate(data_list_base):
            mix_data.append(self.mix_format % (data_value, data_list_use1[data_key], data_list_use2[data_key]))
        print(mix_data)
        return mix_data


class DataWrite(BaseUse):
    def __init__(self, filename, keyword, mix_data=None):
        BaseUse.__init__(self, filename, keyword)
        self.mix_data = mix_data

    def write_data(self):
        Nomax = len(self.mix_data)
        for row in self.ws.iter_rows():
            for cell in row:
                if cell.value == self.keyword:
                    for No in range(Nomax):
                        self.ws['%s%d' % (cell.column, cell.row + No + 1)].value = self.mix_data[No]
        self.wb.save(self.filename)


tool = BaseUse(filename='old.xlsx', keyword='tool')
tool_data = tool.old_data_catch()
series = BaseUse(filename='old.xlsx', keyword='series')
series_data = series.old_data_catch()
mix_base = EDataMix(filename='old.xlsx', mix_format=r'%d/%s')
mix_data = mix_base.mix_2(tool_data, series_data)
deal = DataWrite(filename='new.xlsx', keyword='mix', mix_data=mix_data)
deal2 = DataWrite(filename='new.xlsx', keyword='tool', mix_data=tool_data)
deal2.write_data()
deal3 = DataWrite(filename='new.xlsx', keyword='series', mix_data=series_data)
deal3.write_data()
print(tool_data)
print(series_data)
print(mix_data)