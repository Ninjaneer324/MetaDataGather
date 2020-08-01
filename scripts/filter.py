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

best_terms = []
with open("best_terms.txt", "r") as file:
    for line in file:
        res = line.strip().replace("*", ".*")
        best_terms.append(res)

good_terms = []
with open("good_terms.txt", "r") as file:
    for line in file:
        res = line.strip().replace("*", ".*")
        good_terms.append(res)

margin_good_terms = []
with open("margin_good_terms.txt", "r") as file:
    for line in file:
        res = line.strip().replace("*", ".*")
        margin_good_terms.append(res)

neutral_terms = []
with open("neutral_terms.txt", "r") as file:
    for line in file:
        res = line.strip().replace("*", ".*")
        neutral_terms.append(res)

margin_bad_terms = []
with open("margin_bad_terms.txt", "r") as file:
    for line in file:
        res = line.strip().replace("*", ".*")
        margin_bad_terms.append(res)

bad_terms = []
with open("bad_terms.txt", "r") as file:
    for line in file:
        res = line.strip().replace("*", ".*")
        bad_terms.append(res)

unpromising_terms = []
with open("unpromising_terms.txt", "r") as file:
    for line in file:
        res = line.strip().replace("*", ".*")
        unpromising_terms.append(res)

periodic_table = {}
periodic_wb = xlrd.open_workbook("Periodic-Table.xlsx")
sheet = periodic_wb.sheet_by_index(0)
for i in range(1, 96):
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

def containsElement(input_str=""):
    for i in periodic_table:
        base = i
        alloy = i
        if 'symbol' in periodic_table[i]:
            base = periodic_table[i]['symbol'] + r"-.*"
            alloy = r"-.*" + periodic_table[i]['symbol']
        
        alloy_names = None
        if 'alloy_names' in periodic_table[i]:
            alloy_names = periodic_table[i]['alloy_names']
        if (i.lower() in input_str.lower()) or search(base, input_str) or search(alloy, input_str) or (alloy_names is not None and any(item in input_str for item in alloy_names)):
            return True
    return False

def totalScore(input_str=""):
    score = 0
    for i in best_terms:
        reg = r""+i
        if search(reg, input_str):
            score += 10
    
    for i in good_terms:
        reg = r""+i
        if search(reg, input_str):
            score += 3

    for i in margin_good_terms:
        reg = r""+i
        if search(reg, input_str):
            score += 1
    
    for i in margin_bad_terms:
        reg = r""+i
        if search(reg, input_str):
            score -= 1

    for i in bad_terms:
        reg = r""+i
        if search(reg, input_str):
            score -= 3
    
    for i in unpromising_terms:
        reg = r""+i
        if search(reg, input_str):
            score -= 10
    
    return score

uniques = {}

for r in range(1, no_element_sheet.nrows):
    title = str(no_element_sheet.cell_value(r, 5))
    abstract = str(no_element_sheet.cell_value(r, 2))
    keywords = str(no_element_sheet.cell_value(r, 6))
    if title.lower() not in uniques and (containsElement(title) or containsElement(abstract)) and ((totalScore(title) + totalScore(abstract) +totalScore(keywords)) > 0):
        content = {}
        content['reference-type'] = no_element_sheet.cell_value(r,0)
        content['record-number'] = no_element_sheet.cell_value(r, 1)
        content['abstract'] = abstract
        content['author'] = no_element_sheet.cell_value(r, 3)
        content['year'] = no_element_sheet.cell_value(r, 4)
        content['title'] = title
        content['keywords'] = keywords
        content['journal'] = no_element_sheet.cell_value(r, 7)
        content['label'] = no_element_sheet.cell_value(r, 8)
        content['lanl-style'] = no_element_sheet.cell_value(r, 9)
        uniques[title.lower()] = content
        

del no_element
fil_sheet.write(0, 0, "Reference Type")
fil_sheet.write(0, 1, "Record Number")
fil_sheet.write(0, 2, "Year")
fil_sheet.write(0, 3, "Author")
fil_sheet.write(0, 4, "Title")
fil_sheet.write(0, 5, "Abstract")
fil_sheet.write(0, 6, "Keywords")
fil_sheet.write(0, 7, "Journal")
fil_sheet.write(0, 8, "Label")
fil_sheet.write(0, 9, "LANL Style")
row = 1
for i in uniques:
    fil_sheet.write(row, 0, uniques[i]['reference-type'])
    fil_sheet.write(row, 1, uniques[i]['record-number'])
    fil_sheet.write(row, 2, uniques[i]['year'])
    fil_sheet.write(row, 3, uniques[i]['author'])
    fil_sheet.write(row, 4, uniques[i]['title'])
    fil_sheet.write(row, 5, uniques[i]['abstract'])
    fil_sheet.write(row, 6, uniques[i]['keywords'])
    fil_sheet.write(row, 7, uniques[i]['journal'])
    fil_sheet.write(row, 8, uniques[i]['label'])
    fil_sheet.write(row, 9, uniques[i]['lanl-style'])
    row += 1
filtered.close()