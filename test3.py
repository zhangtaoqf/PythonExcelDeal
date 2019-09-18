import xlutils
from xlrd import open_workbook
from xlutils.copy import copy
rb = open_workbook('Excel_test.xls')
rs = rb.sheet_by_index(0)
wb = copy(rb)
ws = wb.get_sheet(0)
ws.write(0, 0, 'changed!')
wb.save('Excel_test.xls')
