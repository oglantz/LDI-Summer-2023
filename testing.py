import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook

# wb = load_workbook('IT LEAP Budget FY23-24 SPRING.xlsx')
#
# ws = wb.active
#
# for row in ws['A1': 'L4']:
#     for cell in row:
#         print(cell.coordinate)


wb = Workbook()
ws = wb.active

lst = [[1,2,3, None, None], [None, None], [1,2,3]]

for i in lst:
    ws.append(i)

wb.save('test.xlsx')
