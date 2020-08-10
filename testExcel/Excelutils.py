from xlutils.copy import copy
import xlrd
import xlwt

oldExcel = xlrd.open_workbook('C:/Users/aby/OneDrive/文档/财务/北京出差pythonTest.xls', formatting_info=True)  # 同时保存格式信息
oldSheet = oldExcel.sheet_by_index(0)

newExcel = copy(oldExcel)
newSheet = newExcel.get_sheet(0)

# newSheet.write(0, 3, 'test')  # 默认格式写入
# newExcel.save('C:/Users/aby/Desktop/newExcel.xls')

newStyle = xlwt.XFStyle()  # 新建空样式（后文代码命名可以不用new）

newFont = xlwt.Font()  # 新建空字体样式
newFont.name = '等线'
newFont.bold = True  # 加粗
newFont.height = 280  # 280 = 14pt * 20
newStyle.font = newFont  # 添加至新建的空样式

newBorders = xlwt.Borders()  # 新建空边框样式
newBorders.top = xlwt.Borders.THIN  # 细上边框
newBorders.bottom = xlwt.Borders.THIN
newBorders.left = xlwt.Borders.THIN
newBorders.right = xlwt.Borders.THIN
newStyle.borders = newBorders  # 添加至新建的空样式

newAlignment = xlwt.Alignment()
newAlignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平居中
newAlignment.vert = xlwt.Alignment.VERT_CENTER  # 垂直居中
newStyle.alignment = newAlignment

newSheet.write(0, 3, 'test', newStyle)
newExcel.save('C:/Users/aby/Desktop/newExcel.xls')