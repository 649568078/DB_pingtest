import xlrd as xd

data =xd.open_workbook ('get_list.xlsx') #打开excel表所在路径
sheet = data.sheet_by_name('测试数据库清单')  #读取数据，以excel表名来打开
d = []
for r in range(sheet.nrows): #将表中数据按行逐步添加到列表中，最后转换为list结构
    d.append(sheet.cell_value(r,2))

print(d)