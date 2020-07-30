import xlsxwriter
import xlrd
import requests
import time
from math import ceil
from re import search
from urllib.parse import unquote

no_element = xlrd.open_workbook("FirstDataBase01.xlsx")
no_element_sheet = no_element.sheet_by_index(0)
filtered = xlsxwriter.Workbook("FirstDataBase02.xlsx")
fil_sheet = filtered.add_worksheet()

uniques = {}

for r in range(1, no_element_sheet.nrows):
    title = str(no_element_sheet.cell_value(r, 2))
    if title.lower() not in uniques:
        content = {}
        content['ids'] = no_element_sheet.cell_value(r, 0)
        content['year'] = no_element_sheet.cell_value(r, 1)
        content['title'] = title
        content['author'] = no_element_sheet.cell_value(r, 3)
        uniques[title.lower()] = content
del no_element
fil_sheet.write(0, 0, "DOI")
fil_sheet.write(0, 1, "Year")
fil_sheet.write(0, 2, "Title")
fil_sheet.write(0, 3, "Author")
row = 1
for i in uniques:
    fil_sheet.write(row, 0, uniques[i]['ids'])
    fil_sheet.write(row, 1, uniques[i]['year'])
    fil_sheet.write(row, 2, uniques[i]['title'])
    fil_sheet.write(row, 3, uniques[i]['author'])
    row += 1
filtered.close()