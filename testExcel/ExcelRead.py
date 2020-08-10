import xlrd

xlsx = xlrd.open_workbook('C:/Users/aby/OneDrive/文档/财务/北京出差.xlsx')
xlsxData = xlsx.sheet_by_index(0)
# xlsxData = xlsx.sheet_by_name('北京出差') # 另一种方法
print(xlsxData.cell_value(1, 2))
print(xlsxData.cell(1, 2).value)  # 别的方法
print(xlsxData.row(1)[2].value)  # 别的方法
