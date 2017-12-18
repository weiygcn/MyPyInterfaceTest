# -*- encoding:utf-8 -*-
from openpyxl import Workbook
wb=Workbook()
sheet = wb.active
sheet.title ="ygw"
sheet['C3']='hello world'
for i in range(10):
    sheet["A%d" % (i+1)].value = i+1
wb.save('newExcel.xlsx')