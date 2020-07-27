import xlsxwriter
import xlrd
import requests
import time
from math import ceil
from re import search
from urllib.parse import unquote

no_element = xlrd.open_workbook("NoElement.xlsx")
no_element_sheet = no_element.sheet_by_index(0)
filtered = xlsxwriter.Workbook("NoElement-filtered.xlsx")
fil_sheet = filtered.add_worksheet()

uniques = {}

for r in range(2, no_element_sheet.nrows):
    title = no_element_sheet.cell_value(r, 2)
    if title.lower() not in uniques:
        content = {}
        content['ids'] = no_element_sheet.cell_value(r, 1)
        content['title'] = title
        content['author'] = no_element_sheet.cell_value(r, 3)
        content['date'] = no_element_sheet.cell_value(r, 4)
        uniques[title.lower()] = content

fil_sheet.write(0, 0, "Query Format")
fil_sheet.write(0, 1, no_element_sheet.cell_value(0, 1))
fil_sheet.write(2, 0, no_element_sheet.cell_value(2, 0))
fil_sheet.write(3, 0, no_element_sheet.cell_value(3, 0))
del no_element
row = 2
for i in uniques:
    fil_sheet.write(row, 1, uniques[i]['ids'])
    fil_sheet.write(row, 2, uniques[i]['title'])
    fil_sheet.write(row, 3, uniques[i]['author'])
    fil_sheet.write(row, 4, uniques[i]['date'])
    row += 1
filtered.close()