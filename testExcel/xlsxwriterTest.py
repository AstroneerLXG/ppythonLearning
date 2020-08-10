import xlsxwriter as xw  # 将xlsxwriter库以xw的名字载入

newWorkbook = xw.Workbook('C:/Users/aby/Desktop/newExcel.xlsx')
sheet0 = newWorkbook.add_worksheet('sheet0')

for i in range(0, 300):
    sheet0.write(0, i, i)

newWorkbook.close()
