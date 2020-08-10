import xlwt

newExcel = xlwt.Workbook()  # 新建工作簿
dataSheet = newExcel.add_sheet('dataSheetTest')  # 在工作簿中新建工作表
dataSheet.write(0, 0, 'test')  # 工作表中的某个单元格内写入
newExcel.save('C:/Users/aby/Desktop/newExcel.xls')  # 保存整个工作簿
