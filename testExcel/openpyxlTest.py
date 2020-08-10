import openpyxl

newWorkbook = openpyxl.load_workbook('C:/Users/aby/OneDrive/文档/财务/北京出差.xlsx')
sheet0 = newWorkbook['Sheet1']  # 导出工作簿中的Sheet1到sheet0变量
sheet0['D1'] = 'Test'  # 在D1格内写入字符串‘Test’
sheet0['D2'] = 233  # 字符串写入为文本格式，此处为数字格式
newWorkbook.save('C:/Users/aby/Desktop/newExcel.xlsx')
