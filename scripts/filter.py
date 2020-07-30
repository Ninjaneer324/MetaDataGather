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
    title = str(no_element_sheet.cell_value(r, 5))
    if title.lower() not in uniques:
        content = {}
        content['reference-type'] = no_element_sheet.cell_value(r,0)
        content['record-number'] = no_element_sheet.cell_value(r, 1)
        content['abstract'] = no_element_sheet.cell_value(r, 2)
        content['author'] = no_element_sheet.cell_value(r, 3)
        content['year'] = no_element_sheet.cell_value(r, 4)
        content['title'] = title
        content['keywords'] = no_element_sheet.cell_value(r, 6)
        content['label'] = no_element_sheet.cell_value(r, 7)
        content['lanl-style'] = no_element_sheet.cell_value(r, 8)
        uniques[title.lower()] = content
del no_element
fil_sheet.write(0, 0, "Reference Type")
fil_sheet.write(0, 1, "Record Number")
fil_sheet.write(0, 2, "Abstract")
fil_sheet.write(0, 3, "Author")
fil_sheet.write(0, 4, "Year")
fil_sheet.write(0, 5, "Title")
fil_sheet.write(0, 6, "Keywords")
fil_sheet.write(0, 7, "Label")
fil_sheet.write(0, 8, "LANL Style")
row = 1
for i in uniques:
    fil_sheet.write(row, 0, uniques[i]['reference-type'])
    fil_sheet.write(row, 1, uniques[i]['record-number'])
    fil_sheet.write(row, 2, uniques[i]['abstract'])
    fil_sheet.write(row, 3, uniques[i]['author'])
    fil_sheet.write(row, 4, uniques[i]['year'])
    fil_sheet.write(row, 5, uniques[i]['title'])
    fil_sheet.write(row, 6, uniques[i]['keywords'])
    fil_sheet.write(row, 7, uniques[i]['label'])
    fil_sheet.write(row, 8, uniques[i]['lanl-style'])
    row += 1
filtered.close()