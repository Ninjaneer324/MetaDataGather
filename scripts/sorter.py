import xlrd
import xlsxwriter

periodic_table = {}
periodic_wb = xlrd.open_workbook("Periodic-Table.xlsx")
sheet = periodic_wb.sheet_by_index(0)
for i in range(1, 104):
    contents = {}
    contents['symbol'] = sheet.cell_value(i, 2)
    contents['pos'] = i
    periodic_table[sheet.cell_value(i, 1)] = contents
del periodic_wb

common_alloys = xlrd.open_workbook("Common Alloys' Names.xlsx")
sheet = common_alloys.sheet_by_index(0)
rows = sheet.nrows
r = 0
while r < rows:
    base = sheet.cell_value(r, 0).strip()
    r += 3
    alloy_names = []
    while r < rows and sheet.cell_value(r, 0) != "-":
        stuff = sheet.cell_value(r, 0)
        first_parenthesis = stuff.find("(")
        second_parenthesis = stuff.find(")")
        alloy_n = stuff[0:second_parenthesis + 1].strip() if (first_parenthesis > -1 and second_parenthesis > -1) else stuff.strip()
        alloy_names.append(alloy_n)
        r += 1
    if base in periodic_table:
        periodic_table[base]['alloy_names'] = alloy_names
    else:
        contents = {'alloy_names':alloy_names}
        periodic_table[base] = contents
    r += 1