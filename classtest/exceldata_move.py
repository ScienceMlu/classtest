#! /usr/bin/python
# -*- coding: utf-8 -*-
from openpyxl import load_workbook


class BaseUse:
    def __init__(self, filename, keyword, mix_format):
        self.filename = filename
        self.keyword = keyword
        self.wb = load_workbook(filename)
        self.ws = self.wb['時間_承認']
        mix_data = []
        self.mix_data = mix_data
        self.mix_format = mix_format

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
    def mix_2(self, data_list_base, data_list_use):
        for data_key, data_value in enumerate(data_list_base):
            self.mix_data.append(self.mix_format % (data_value, data_list_use[data_key]))
        print(self.mix_data)
        return self.mix_data

    def mix_3(self, data_list_base, data_list_use1, data_list_use2):
        for data_key, data_value in enumerate(data_list_base):
            self.mix_data.append(self.mix_format % (data_value, data_list_use1[data_key], data_list_use2[data_key]))
        print(self.mix_data)
        return self.mix_data


class DataWrite(EDataMix):
    def write_data(self):
        Nomax = len(self.mix_data)
        for row in self.ws.iter_rows():
            for cell in row:
                if cell.value == self.keyword:
                    for No in range(Nomax):
                        self.ws['%s%d' % (cell.column, cell.row + No + 1)].value = self.mix_data[No]
        self.wb.save(self.filename)
