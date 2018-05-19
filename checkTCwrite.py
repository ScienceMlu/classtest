#! /usr/bin/python
# -*- coding: utf-8 -*-


import openpyxl

'''
filename=input('输入检查文件名：')
wb=openpyxl.load_workbook(filename)
'''
# 检查sheet内容
sheetName = []
group = []


def check_tc(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet_name = wb.sheetnames
    group_data = []
    for j in range(len(sheet_name)):
        sheet1 = wb[sheet_name[j]]
        listA = []
        listB = []
        listC = []
        listD = []
        listE = []
        areaA = []
        areaB = []
        areaC = []
        areaD = []
        areaE = []
        countNumber = []
        for row in sheet1.iter_rows():
            for cell in row:
                if str(cell.fill.fgColor.rgb) == 'FFFFFF00':
                    if cell.value is None:
                        data_a = [cell.row, cell.column]
                        areaA.append(cell.coordinate)
                        listA.append(data_a)
                    elif cell.value == '〇':
                        print('ok')
                    elif str(cell.value).find('×') != -1:
                        data_b = [cell.row, cell.column]
                        areaB.append(cell.coordinate)
                        listB.append(data_b)
                    elif cell.value == '△':
                        data_c = [cell.row, cell.column]
                        areaC.append(cell.coordinate)
                        listC.append(data_c)
                    elif str(cell.value).find('N/A') != -1:
                        data_d = [cell.row, cell.column]
                        listD.append(data_d)
                        areaD.append(cell.coordinate)
                    else:
                        data_e = [cell.row, cell.column]
                        listE.append(data_e)
                        areaE.append(cell.coordinate)
        countNumber.append(len(areaA))
        countNumber.append(len(areaB))
        countNumber.append(len(areaC))
        countNumber.append(len(areaD))
        countNumber.append(len(areaE))
        group_data.append([areaA, areaB, areaC, areaD, areaE, countNumber])
        # group= dict.fromkeys(sheetName_now,[listA,listB,listC,listD,listE,areaA,areaB,areaC,areaD,areaE,countNumber])
        '''
        print(areaA)
        print(areaB)
        print(areaC)
        print(areaD)
        print(areaE)
        print(countNumber)
        print('------------')
        '''
    return group_data, sheet_name


# 写结果
def write_result(group_data, sheet_name, save_name):
    from openpyxl.styles import Alignment
    from openpyxl import Workbook
    wb = Workbook()

    for k in range(len(sheet_name)):
        ws = wb.create_sheet(index=k)

        '''设置字体大小颜色，单元格背景'''

        # 合并单元格'
        ws.merge_cells('A1:B2')
        # 空值
        ws.merge_cells('C1:D2')
        # △
        ws.merge_cells('E1:F2')
        # ×
        ws.merge_cells('G1:H2')
        # N/A
        ws.merge_cells('I1:J2')
        # 格式不符

        # 居中单元格
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws['C1'].alignment = Alignment(horizontal='center', vertical='center')
        ws['E1'].alignment = Alignment(horizontal='center', vertical='center')
        ws['G1'].alignment = Alignment(horizontal='center', vertical='center')
        ws['I1'].alignment = Alignment(horizontal='center', vertical='center')

        # 填写格式
        ws['A1'].value = '空值'
        ws['C1'].value = '△'
        ws['E1'].value = '×'
        ws['G1'].value = 'N/A'
        ws['I1'].value = '格式不符'

        i = 0
        for cols in range(1, 10, 2):
            ws.cell(row=3, column=cols, value='个数')
        for cols in range(2, 12, 2):
            ws.cell(row=3, column=cols, value='位置')

        # 写入位置数据
        # 空值位置
        data_len = len(group_data[k][0]) + 4
        print(data_len - 4)
        for rows in range(4, data_len):
            ws.cell(column=2, row=rows, value=group_data[k][0][i])
            if i > data_len - 4:
                break
            i = i + 1
        i = 0

        # △位置

        data_len = len(group_data[k][1]) + 4
        print(data_len - 4)
        for rows in range(4, data_len):
            ws.cell(column=4, row=rows, value=group_data[k][1][i])
            if i > data_len - 4:
                break
            i = i + 1
        i = 0

        # ×位置

        data_len = len(group_data[k][2]) + 4
        print(data_len - 4)
        for rows in range(4, data_len):
            ws.cell(column=6, row=rows, value=group_data[k][2][i])
            if i > data_len - 4:
                break
            i = i + 1
        i = 0

        # N/A位置

        data_len = len(group_data[k][3]) + 4
        print(data_len - 4)
        for rows in range(4, data_len):
            ws.cell(column=8, row=rows, value=group_data[k][3][i])
            if i > data_len - 4:
                break
            i = i + 1
        i = 0

        # 格式不符合位置

        data_len = len(group_data[k][4]) + 4
        print(data_len - 4)
        for rows in range(4, data_len):
            ws.cell(column=10, row=rows, value=group_data[k][4][i])
            if i > data_len - 4:
                break
            i = i + 1
        i = 0
        # 个数
        data_len = len(group_data[k][5])
        print(data_len)
        for cols in range(1, 10, 2):
            ws.cell(row=4, column=cols, value=group_data[k][5][i])
            # areaA:0,areaB:1,areaC:2,areaD:3,areaE:4,countNumber:5
            i += 1
            if i > data_len:
                break
        print('================')

        # 保存

    wb.save(save_name)


'''                
print('Wait a moment.....')



t=open('coordinate.txt','w')
if(listA!=[]):
    t.write('空值='+pprint.pformat(listA) + '\n' +'个数='+str(len(listA))+'\n'+'========================='+'\n')
if(listB!=[]):
    t.write('×所在line=' + pprint.pformat(listB)  + '\n' +'个数='+str(len(listB))+ '\n' + '====================='+'\n')
if(listC!=[]):
    t.write('△所在line=' + pprint.pformat(listC)  + '\n' +'个数='+str(len(listC))+ '\n' + '==================='+'\n')
if(listD!=[]):
    t.write('N/A所在line='+pprint.pformat(listD)  + '\n' +'个数='+str(len(listD))+ '\n' + '====================='+'\n')
if(listE!=[]):
    t.write('书写不规范所在line=' + pprint.pformat(listE) + '\n' +'个数='+str(len(listE))+ '\n'+'==============='+'\n')
t.close()
print('OK,Thank you for using!!!!')
'''
