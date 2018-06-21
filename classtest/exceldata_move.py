#! /usr/bin/python
# -*- coding: utf-8 -*-
import os

from openpyxl import load_workbook, Workbook


class Deal(object):
    def __init__(self, dom_path):
        self.dom_path = dom_path
        i = 0
        if not os.path.exists('end%d' % i):
            os.mkdir('end%d' % i)
            out_path = 'end%d' % i
        else:
            i += 1
            os.mkdir('end%d' % i)
            out_path = 'end%d' % i
        self.out_path = r'%s\%s' % (out_path, dom_path)

    def find_files(self):
        data_excel = []
        old_data_path = r'%s\old' % self.dom_path
        os.chdir(old_data_path)
        files = os.listdir(old_data_path)
        for F in files:
            if os.path.splitext(F)[-1] == '.xlsx':
                data_excel.append(F)
        return data_excel


class BaseUse(object):
    def __init__(self, filename, sheetname):
        keyword_old = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10']
        keyword_new = ['I', 'II', 'III', 'IV', 'V', 'VI']
        self.keyword_old = keyword_old
        self.keyword_new = keyword_new
        self.filename = filename
        self.wb = load_workbook(filename)
        self.ws = self.wb[sheetname]

    def old_data_catch(self, keyword):
        for u_row in self.ws.iter_rows():
            for cell in u_row:
                if cell.value == keyword:
                    keyRow = cell.row
        data_list = []
        Nomax = self.ws.max_row - keyRow
        for c_row in self.ws.iter_rows():
            for cell in c_row:
                if cell.value == keyword:
                    for No in range(Nomax):
                        data_list.append(self.ws['%s%d' % (cell.column, cell.row + No + 1)].value)
        return data_list

    def mix_2(self, data_list_base, data_list_use, mix_format):
        mix_data = []
        for data_key, data_value in enumerate(data_list_base):
            mix_data.append(mix_format % (data_value, data_list_use[data_key]))
        print(mix_data)
        return mix_data

    def mix_3(self, data_list_base, data_list_use1, data_list_use2, mix_format):
        mix_data = []
        for data_key, data_value in enumerate(data_list_base):
            mix_data.append(mix_format % (data_value, data_list_use1[data_key], data_list_use2[data_key]))
        print(mix_data)
        return mix_data

    def data_save(self):
        old_multi_data = []
        first_D = self.old_data_catch(keyword=self.keyword_old[0])
        old_multi_data.append(first_D)
        second_D = self.old_data_catch(keyword=self.keyword_old[1])
        old_multi_data.append(second_D)
        third_D = self.old_data_catch(keyword=self.keyword_old[2])
        old_multi_data.append(third_D)
        forth_D = self.old_data_catch(keyword=self.keyword_old[3])
        old_multi_data.append(forth_D)
        fifth_D = self.old_data_catch(keyword=self.keyword_old[4])
        old_multi_data.append(fifth_D)
        sixth_D = self.old_data_catch(keyword=self.keyword_old[5])
        old_multi_data.append(sixth_D)
        seventh_D = self.old_data_catch(keyword=self.keyword_old[6])
        old_multi_data.append(seventh_D)
        eight_D = self.old_data_catch(keyword=self.keyword_old[7])
        old_multi_data.append(eight_D)
        ninth_D = self.old_data_catch(keyword=self.keyword_old[8])
        old_multi_data.append(ninth_D)
        ten_D = self.old_data_catch(keyword=self.keyword_old[9])
        old_multi_data.append(ten_D)
        return old_multi_data
        # =========


class DataWrite(object):
    def __init__(self, deal_path, old_name, keyword, data):
        os.chdir(deal_path)
        self.keyword = keyword
        self.data = data
        wb = Workbook()
        ws = wb.create_sheet('output')
        self.wb = wb
        self.ws = ws
        self.filename = 'NEW_%s' % old_name
        wb.save(self.filename)

    def cell_format(self):
        for row in self.ws.max_row:
            self.ws.row_dimensions[row].height = 83

    def write_data(self):
        Nomax = len(self.data)
        for row in self.ws.iter_rows():
            for cell in row:
                if cell.value == self.keyword:
                    for No in range(Nomax):
                        self.ws['%s%d' % (cell.column, cell.row + No + 1)].value = self.data[No]
        self.wb.save(self.filename)


def use():
    t = Deal(dom_path=input('输入处理文件夹路径：'))
    file_list = t.find_files()
    for F in file_list:
        old_deal = BaseUse(filename=F, sheetname='ok')
        old_data_list1 = old_deal.data_save()
        old_data_mix2 = old_deal.mix_2(old_data_list1[0], old_data_list1[1], mix_format=r'[%s]%s')
        old_data_mix2_1 = old_deal.mix_2(old_data_list1[3], old_data_list1[4], mix_format=r'%s：%s')
        old_data_mix3 = old_deal.mix_3(old_data_list1[7], old_data_list1[8], old_data_list1[9], mix_format=r'%s/%s/%s')
        D1 = DataWrite(deal_path=t.out_path, old_name=F, keyword=old_deal.keyword_new[0], data=old_data_mix2)
        D1.write_data()
        D2 = DataWrite(deal_path=t.out_path, old_name=F, keyword=old_deal.keyword_new[1], data=old_data_list1[2])
        D2.write_data()
        D3 = DataWrite(deal_path=t.out_path, old_name=F, keyword=old_deal.keyword_new[2], data=old_data_mix2_1)
        D3.write_data()
        D4 = DataWrite(deal_path=t.out_path, old_name=F, keyword=old_deal.keyword_new[3], data=old_data_mix3)
        D4.write_data()
        D5 = DataWrite(deal_path=t.out_path, old_name=F, keyword=old_deal.keyword_new[4], data=old_data_list1[5])
        D5.write_data()
        D6 = DataWrite(deal_path=t.out_path, old_name=F, keyword=old_deal.keyword_new[5], data=old_data_list1[6])
        D6.write_data()

if __name__ == '__main__':
    '''
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
    '''

    use()