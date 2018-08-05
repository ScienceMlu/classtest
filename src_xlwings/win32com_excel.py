from win32com.client import Dispatch
import win32com.client


class easyExcel:
    """A utility to make it easier to get at Excel.    Remembering
    to save the data is your problem, as is    error handling.
    Operates on one workbook at a time."""

    def __init__(self, filename=None):  # 打开文件或者新建文件（如果不存在的话）
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    def save(self, newfilename=None):  # 保存文件
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()

    def close(self):  # 关闭文件
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp

    def getCell(self, sheet, row, col):  # 获取单元格的数据
        "Get value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value

    def setCell(self, sheet, row, col, value):  # 设置单元格的数据
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value

    def setCellformat(self, sheet, row, col):  # 设置单元格的数据
        "set value of one cell"
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Font.Size = 15  # 字体大小
        sht.Cells(row, col).Font.Bold = True  # 是否黑体
        sht.Cells(row, col).Name = "Arial"  # 字体类型
        sht.Cells(row, col).Interior.ColorIndex = 3  # 表格背景
        # sht.Range("A1").Borders.LineStyle = xlDouble
        sht.Cells(row, col).BorderAround(1, 4)  # 表格边框
        sht.Rows(3).RowHeight = 30  # 行高
        sht.Cells(row, col).HorizontalAlignment = -4131  # 水平居中xlCenter
        sht.Cells(row, col).VerticalAlignment = -4160  #

    def deleteRow(self, sheet, row):
        sht = self.xlBook.Worksheets(sheet)
        sht.Rows(row).Delete()  # 删除行
        sht.Columns(row).Delete()  # 删除列

    def getRange(self, sheet, row1, col1, row2, col2):  # 获得一块区域的数据，返回为一个二维元组
        "return a 2d array (i.e. tuple of tuples)"
        sht = self.xlBook.Worksheets(sheet)
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value

    def addPicture(self, sheet, pictureName, Left, Top, Width, Height):  # 插入图片
        "Insert a picture in sheet"
        sht = self.xlBook.Worksheets(sheet)
        sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)

    def cpSheet(self):  # 复制工作表
        "copy sheet"
        shts = self.xlBook.Worksheets
        shts(1).Copy(None, shts(1))

    def inserRow(self, sheet, row):
        sht = self.xlBook.Worksheets(sheet)
        sht.Rows(row).Insert(1)

    # 下面是一些测试代码。


import win32com

APP_TYPE = 'Excel.Application'

xlBlack, xlRed, xlGray, xlBlue = 1, 3, 15, 41
xlBreakFull = 1

# 初始化应用程序
xls = win32com.client.Dispatch(APP_TYPE)
xls.Visible = True
book = xls.Workbooks.Add()
sheet = book.Worksheets(1)

# 插入标题
ROW_PER_PAGE, COL_PER_ROW = 10, 10
row_index, col_index = 1
title_range = sheet.Range(sheet.Cells(row_index, col_index), sheet.Cells(row_index, COL_PER_ROW))
title_range.MergeCells = True
title_range.Font.Bold, title_range.Interior.ColorIndex = True, xlGray
title_range.Value = 'Hello,Word'
row_index += 1

# 插入内容 10*10的数列
for row in range(0, 10):
    col_index = 1
    for col in range(0, COL_PER_ROW):
        cell_range = sheet.Cells(row_index, col_index)
        cell_range.Font.Color, cell_range.Value = xlBlue, str(row)
    row_index += 1

# 插入分页符
right_bottom_range = sheet.Cells(row_index, COL_PER_ROW + 1)
right_bottom_range.PageBreak = xlBreakFull

# 插入图片
col_index = 1
lt_range = sheet.Cells(row_index, col_index)
graph_width = sheet.Range(sheet.Cells(row_index, 1), sheet.Cells(row_index, COL_PER_ROW)).Width
graph_height = sheet.Range(sheet.CellS(row_index, 1), sheet.Cells(row_index + ROW_PER_PAGE, 1)).Height
sheet.Shapes.AddPicture('C:\\test.jpg', False, True, lt_range.Left, lt_range.Top, graph_width, graph_height)

if __name__ == "__main__":
    # PNFILE = r'c:/screenshot.bmp'
    xls = easyExcel(r'E:\java\python\Test\src_xlwings\a.xlsm')
    # xls.addPicture('Sheet1', PNFILE, 20,20,1000,1000)
    # xls.cpSheet('Sheet1')
    win32com.client.Dispatch('Excel.Application').Workbooks.Open(r'E:\java\python\Test\src_xlwings\a.xlsm')
    sht1 = xls.xlBook.Worksheets('DataMap')
    sht1.Cells(8,8).BorderAround(1, 4)
    print("*******beginsetCellformat********")
    # while(row<5):
    #   while(col<5):
    #       xls.setCellformat('sheet1',row,col)
    #       col += 1
    #       print("row=%s,col=%s" %(row,col))
    #   row += 1
    #   col=1
    #   print("*******row********")
    # print("*******endsetCellformat********")
    # print("*******deleteRow********")
    # xls.deleteRow('sheet1',5)
    xls.save()
    xls.close()


