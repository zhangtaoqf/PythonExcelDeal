import xlrd
wb = xlrd.open_workbook('Excel_test.xls')
print(wb.sheet_names())
wb.get_sheets()
