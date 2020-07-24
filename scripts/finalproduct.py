import xlrd
import xlsxwriter
from re import search

read = xlrd.open_workbook("NoElement.xlsx")
read_sheet = read.sheet_by_index(0)

periodic_wb = xlrd.open_workbook("Periodic-Table.xlsx")
periodic_sheet = periodic_wb.sheet_by_index(0)

write = xlsxwriter.Workbook("FinalProduct.xlsx")
worksheet = write.add_worksheet()
for i in range(1, 96):
    stuff = periodic_sheet.cell_value(i, 2)
    worksheet.write(0, i, stuff)
    worksheet.write(i, 0, stuff)

periodic_table = {}
for i in range(1, 96):
    contents = {}
    contents['name'] = periodic_sheet.cell_value(i, 1)
    contents['alloys'] = periodic_sheet.cell_value(i,3)
    contents['pos'] = i
    periodic_table[periodic_sheet.cell_value(i, 2)] = contents
del periodic_wb

n_rows = int(read_sheet.nrows)
if read_sheet.cell_value(3, 0) > 0:
    for i in periodic_table:
        for j in periofic_table:
            for r in range(2, n_rows):
                #to handle base-alloy format
                dont_want_alloy = r"-.*"+i
                dont_want_base = j+r"-.*"

                base = i + r"-.*"
                alloy = r"-.*" + j

                title = read_sheet.cell_value(r, 2)
                if (not (search(dont_want_alloy, title) or search(dont_want_base, title))) and search(base, title) and search(alloy, title):
                    print("Found")