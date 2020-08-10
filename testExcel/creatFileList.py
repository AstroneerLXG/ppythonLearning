import os
import xlwt

fileDir = 'C:/Users/aby/OneDrive/文档/财务/'

newWorkbook = xlwt.Workbook()
newSheet = newWorkbook.add_sheet('sheetTest')

n = 0
for i in os.listdir(fileDir):  # listdir用于获取目录下文件清单
    newSheet.write(n, 0, i)
    n += 1

newWorkbook.save('C:/Users/aby/Desktop/fileList.xls')