import xlrd
import xlsxwriter
import re
from re import search
import time

def is_all_caps(s):
    return all(char.isupper() for char in s)

periodic_table = {}
periodic_wb = xlrd.open_workbook("Periodic-Table.xlsx")
sheet = periodic_wb.sheet_by_index(0)
for i in range(1, 107):
    contents = {}
    contents['symbol'] = sheet.cell_value(i, 2)
    contents['pos'] = i
    periodic_table[sheet.cell_value(i, 1)] = contents

nxn_table = xlsxwriter.Workbook('FinalProduct.xlsx')
worksheet = nxn_table.add_worksheet()
for i in range(1, 107):
    stuff = sheet.cell_value(i, 2)
    worksheet.write(0, i, stuff)
    worksheet.write(i, 0, stuff)
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
del common_alloys
def checkPure(input_str, elem):
    #print(periodic_table[elem]['symbol']+"-"+periodic_table[elem]['symbol'])
    return True if search(r"pure " + elem, input_str, re.IGNORECASE) else False

def checkDashNotation(input_str, base, alloy):
    if periodic_table[base]['pos'] < 104 and periodic_table[alloy]['pos'] < 104:
        #print(periodic_table[base]['symbol']+"-"+periodic_table[alloy]['symbol'])
        reg_base = r"[^-]+\s+"+periodic_table[base]['symbol']+r"-.*" + periodic_table[alloy]['symbol']
        reg_base_2 = r"[^-]+\s+"+base+r"-.*" +alloy
        return True if (search(reg_base, input_str, re.IGNORECASE) or search(reg_base_2, input_str, re.IGNORECASE)) else False
    return False

def checkNoDash(input_str, base,alloy):
    if periodic_table[base]['pos'] < 104 and periodic_table[alloy]['pos'] < 104:
        #print(periodic_table[base]['symbol']+"-"+periodic_table[alloy]['symbol'])
        reg_base = r"" + periodic_table[base]['symbol'] + r"[a-zA-Z]*" + periodic_table[alloy]['symbol']
        reg_base_2 = r"" + periodic_table[base]['symbol'] + periodic_table[alloy]['symbol']
        return True if search(reg_base, input_str) or search(reg_base_2, input_str) else False
    return False

def capitalizeWords(input_str):
    s = ""
    for i in input_str.split(" "):
        s += i.capitalize() + " "
    return s

def checkAlloyNames(input_str, base):
    pairs = []
    if 'alloy_names' in periodic_table[base]:
        alloy_names = periodic_table[base]['alloy_names']
        for a in alloy_names:
            alloying_low = a[a.find("(") + 1: a.find(")")].split(", ")
            alloying = [capitalizeWords(j) for j in alloying_low]
            if (search(r""+a[0:a.find("(")].strip(), input_str) if is_all_caps(a[0:a.find("(")]) else search(r""+a[0:a.find("(")].strip(), input_str, re.IGNORECASE)):
                for j in alloying:
                    pairs.append([periodic_table[base]['pos'], periodic_table[j]['pos']])
    return pairs
unadded = []

filtered = xlrd.open_workbook('FirstDataBase02v2.xlsx')
fil_sheet = filtered.sheet_by_index(0)
rows = fil_sheet.nrows
for r in range(1, rows):
    title = fil_sheet.cell_value(r, 4)
    abstract = fil_sheet.cell_value(r, 5)
    keywords = fil_sheet.cell_value(r, 6)
    checks = [title, abstract, keywords]
    year = int(fil_sheet.cell_value(r, 2))
    author_f = fil_sheet.cell_value(r, 3)
    author = author_f[0:author_f.find(",")].strip().lower()
    for b in periodic_table:
        for a in periodic_table:
            print(b+"-"+a)
            if b == a:
                if checkPure(title, b) or checkPure(abstract, b) or checkPure(keywords, b):
                    pos = periodic_table[b]['pos']
                    worksheet.write(pos, pos, str(year)+author)
                    continue
            if checkDashNotation(title, b, a) or checkDashNotation(abstract, b, a) or checkDashNotation(keywords, b, a):
                pos_b = periodic_table[b]['pos']
                pos_a = periodic_table[a]['pos']
                worksheet.write(pos_b, pos_a, str(year)+author)
            if checkNoDash(title, b, a) or checkNoDash(abstract, b, a) or checkNoDash(keywords, b, a):
                pos_b = periodic_table[b]['pos']
                pos_a = periodic_table[a]['pos']
                worksheet.write(pos_b, pos_a, str(year)+author)
    
        alloy_names_temp = checkAlloyNames(title, b) + checkAlloyNames(abstract, b) + checkAlloyNames(keywords, b)
        if alloy_names_temp:
            for p in alloy_names_temp:
                worksheet.write(p[0], p[1], str(year)+author)

del filtered
nxn_table.close()