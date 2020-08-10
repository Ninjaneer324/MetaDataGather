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
        reg_base = r"[^-]+\s+"+periodic_table[base]['symbol']+r"-.*" + periodic_table[alloy]['symbol']+r"[-\s]+"
        reg_base_2 = r"[^-]+\s+"+base+r"-.*" +alloy
        return True if (search(reg_base, input_str, re.IGNORECASE) or search(reg_base_2, input_str, re.IGNORECASE)) else False
    return False

def checkNoDash(input_str, base,alloy):
    if periodic_table[base]['pos'] < 104 and periodic_table[alloy]['pos'] < 104:
        #print(periodic_table[base]['symbol']+"-"+periodic_table[alloy]['symbol'])
        reg_base = r"" + periodic_table[base]['symbol'] + r"[a-zA-Z]" + periodic_table[alloy]['symbol']
        reg_base_2 = r"" + periodic_table[base]['symbol'] + periodic_table[alloy]['symbol']
        return True if search(reg_base, input_str) or search(reg_base_2, input_str) else False
    return False

def capitalizeWords(input_str):
    s = ""
    for i in input_str.split(" "):
        s += i.capitalize() + " "
    return s

def checkAlloyNames(input_str):
    pairs = []
    for base in periodic_table:
        if 'alloy_names' in periodic_table[base]:
            alloy_names = periodic_table[base]['alloy_names']
            for a in alloy_names:
                alloying_low = a[a.find("(") + 1: a.find(")")].split(", ")
                alloying = [capitalizeWords(j).strip() for j in alloying_low]
                if (search(r""+a[0:a.find("(")].strip(), input_str) if is_all_caps(a[0:a.find("(")]) else search(r""+a[0:a.find("(")].strip(), input_str, re.IGNORECASE)):
                    for j in alloying:
                        pairs.append([periodic_table[base]['pos'], periodic_table[j]['pos']])
    return pairs
unadded = []

periodic_array = [["" for i in range(106)] for j in range(106)]

filtered = xlrd.open_workbook('FirstDataBase02v2.xlsx')
fil_sheet = filtered.sheet_by_index(0)
start = 1
records_counted = []
for r in range(start, fil_sheet.nrows):
    counted = False
    title = fil_sheet.cell_value(r, 4)
    abstract = fil_sheet.cell_value(r, 5)
    keywords = fil_sheet.cell_value(r, 6)
    checks = [title, abstract, keywords]
    year = int(fil_sheet.cell_value(r, 2))
    author_f = fil_sheet.cell_value(r, 3)
    author = author_f[0:author_f.find(",")].strip().lower()
    alloy_names_temp = checkAlloyNames(title) + checkAlloyNames(abstract) + checkAlloyNames(keywords)
    if alloy_names_temp:
        print("Named alloys found")
        counted = True
        for p in alloy_names_temp:
            
                periodic_array[p[0] - 1][p[1] - 1] = periodic_array[p[0] - 1][p[1] - 1] + (";" if periodic_array[p[0] - 1][p[1] - 1] else "") +str(year)+author
            
    for b in periodic_table:       
        if checkPure(title, b) or checkPure(abstract, b) or checkPure(keywords, b):
                counted = True
                pos = periodic_table[b]['pos']
                periodic_array[pos - 1][pos - 1] = periodic_array[pos - 1][pos - 1] + (";" if periodic_array[pos - 1][pos - 1] else "") +str(year)+author
        '''if checkDashNotation(title, b, a) or checkDashNotation(abstract, b, a) or checkDashNotation(keywords, b, a):
            print(fil_sheet.cell_value(r, 1))
            print("checkDashNotation")
            pos_b = periodic_table[b]['pos']
            pos_a = periodic_table[a]['pos']
            periodic_array[pos_b - 1][pos_a - 1] = periodic_array[pos_b - 1][pos_a - 1] + (";" if periodic_array[pos_b - 1][pos_a - 1] else "") +str(year)+author
            worksheet.write(pos_b, pos_a, str(year)+author)
        elif checkNoDash(title, b, a) or checkNoDash(abstract, b, a) or checkNoDash(keywords, b, a):
            print(fil_sheet.cell_value(r, 1))
            print("checkNoDashNotation")
            pos_b = periodic_table[b]['pos']
            pos_a = periodic_table[a]['pos']
            periodic_array[pos_b - 1][pos_a - 1] = periodic_array[pos_b - 1][pos_a - 1] + (";" if periodic_array[pos_b - 1][pos_a - 1] else "") +str(year)+author
            worksheet.write(pos_b, pos_a, str(year)+author)'''
    if counted:
        records_counted.append(int(fil_sheet.cell_value(r, 1)))
print("Left:", fil_sheet.nrows - len(records_counted) - 1)

'''for r in range(1, fil_sheet.nrows):
    if int(fil_sheet.cell_value(r, 1)) not in records_counted:
        print("Title:", fil_sheet.cell_value(r, 4), "\n")
        print("Abstract:", fil_sheet.cell_value(r, 5), "\n")
        print("Keywords:", fil_sheet.cell_value(r, 6), "\n")
        
        base = input("Base Alloy? (separate with comma): ").strip()
        alloy = input("Alloy elements? (separate with comma, then base with semicolon): ").strip()

        if base.lower() != "n/a" and alloy.lower() != "n/a":
            base_elems = re.split(r",\s*", base)
            alloy_elems = re.split(r";\s*", alloy)
            for b in range(len(base_elems)):
                a_mini = re.split(r',\s*', alloy_elems[b])
                for a in a_mini:
                    author_f = fil_sheet.cell_value(r, 3)
                    author = author_f[0:author_f.find(",")].strip().lower()
                    periodic_array[periodic_table[base_elems[b].capitalize()]['pos']][periodic_table[a.capitalize()]['pos']] += str(int(fil_sheet.cell_value(r,2))) + author'''

for base in range(len(periodic_array)):
    for alloy in range(len(periodic_array[base])):
        worksheet.write(base + 1, alloy + 1, periodic_array[base][alloy])

del filtered
nxn_table.close()