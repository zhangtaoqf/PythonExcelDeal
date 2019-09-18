# -*- coding:UTF-8 -*-
# 头部一定要指定coding类型，要不然会编码出错

# description:本项目是一个使用xlrd的demo，从元数据中，合并了两列数据

# 打开一个已经存在的项目
import xlrd
# 写一个项目
import xlwt

# 创建一个worksheet
workbook = xlwt.Workbook(encoding='utf-8')
sheet1 = workbook.add_sheet('result')
# 打开一个已经存在的文件如：
openWb = xlrd.open_workbook('Excel_test.xls')
# 获取第一个sheet表
openSheet = openWb.sheet_by_index(0)
# 获取列数
cols = openSheet.ncols
# 获取行数
rows = openSheet.nrows

for row in range(0, rows - 1):
    for col in range(0, cols - 1):
        if row == 0:
            sheet1.write(row, col, "列%d" % (col))
        elif col == 1:
            sheet1.write(row, col, ("%d行%d列" % (row, col)))
        else:
            sheet1.write(row, col, openSheet.cell_value(row, col))

workbook.save('test1.xls')
# 遍历第一个表，
# for row in range(0,rows-1):
#     if row != 0:
#         sheet1.write(row, 1, openSheet.cell_value(row, 1) + openSheet.cell_value(row, 2))
#     else:
#         sheet1.write(row, 1, "城市")
#     sheet1.write(row, 0, openSheet.cell_value(row, 0))
#     # 遍历剩下的行数据
#     for col in range(2, cols -1):
#         sheet1.write(row, col,openSheet.cell_value(row, col + 1))
# 这个获取数据
# 保存

# 获取指定title的列

# 获取单元格内容

# 设置一个单元格内容


