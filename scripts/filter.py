import xlsxwriter
import xlrd
import requests
import time
from math import ceil
import re
from re import search
from re import findall
from urllib.parse import unquote

no_element = xlrd.open_workbook("WOSMasterList.xlsx")
no_element_sheet = no_element.sheet_by_index(0)
filtered = xlsxwriter.Workbook("WOSMasterList-withscore.xlsx")
fil_sheet = filtered.add_worksheet()

def is_all_caps(s):
    return all(char.isupper() for char in s)

'''def allTerms(listOfTerms = [], input_str = "", delimiter = " "):
    words = []
    for i in listOfTerms:
        r = r"" + i
        if not is_all_caps(i):
            words += findall(r, input_str, re.IGNORECASE)
        else:
            words += findall(r, input_str)
    return words'''

def allTerms(listOfTerms = [], input_str = "", delimiter = " "):
    words = {}
    for i in listOfTerms:
        r = r"" + i
        if not is_all_caps(i):
            if search(r, input_str, re.IGNORECASE):
                words[i] = findall(r, input_str, re.IGNORECASE)
        else:
            if search(r, input_str):
                words[i] = findall(r, input_str)
    return words

best_terms = []
with open("best_terms.txt", "r") as file:
    for line in file:
        white_space_before = "\s" if line.strip().startswith("*") else ""
        white_space_after = "\s"
        res = white_space_before + line.strip().replace("*", ".*") + white_space_after
        best_terms.append(res)

good_terms = []
with open("good_terms.txt", "r") as file:
    for line in file:
        white_space_before = "\s" if line.strip().startswith("*") else ""
        white_space_after = "\s"
        res = white_space_before + line.strip().replace("*", ".*") + white_space_after
        good_terms.append(res)

margin_good_terms = []
with open("margin_good_terms.txt", "r") as file:
    for line in file:
        white_space_before = "\s" if line.strip().startswith("*") else ""
        white_space_after = "\s"
        res = white_space_before + line.strip().replace("*", ".*") + white_space_after
        margin_good_terms.append(res)

neutral_terms = []
with open("neutral_terms.txt", "r") as file:
    for line in file:
        white_space_before = "\s" if line.strip().startswith("*") else ""
        white_space_after = "\s"
        res = white_space_before + line.strip().replace("*", ".*") + white_space_after
        neutral_terms.append(res)

margin_bad_terms = []
with open("margin_bad_terms.txt", "r") as file:
    for line in file:
        white_space_before = "\s" if line.strip().startswith("*") else ""
        white_space_after = "\s"
        res = white_space_before + line.strip().replace("*", ".*") + white_space_after
        margin_bad_terms.append(res)

bad_terms = []
with open("bad_terms.txt", "r") as file:
    for line in file:
        white_space_before = "\s" if line.strip().startswith("*") else ""
        white_space_after = "\s"
        res = white_space_before + line.strip().replace("*", ".*") + white_space_after
        bad_terms.append(res)

unpromising_terms = []
with open("unpromising_terms.txt", "r") as file:
    for line in file:
        white_space_before = "\s" if line.strip().startswith("*") else ""
        white_space_after = "\s"
        res = white_space_before + line.strip().replace("*", ".*") + white_space_after
        unpromising_terms.append(res)

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
        if (i.lower() in input_str.lower()) or search(base, input_str) or search(alloy, input_str) or (alloy_names is not None and any((search(r""+(item[0:item.find("(") - 1] if item.find("(") > -1 else item), input_str) if is_all_caps(item[0:item.find("(") - 1]) else search(r"" + (item[0:item.find("(") - 1] if item.find("(") > -1 else item), input_str, re.IGNORECASE)) for item in alloy_names)):
            return True
    return False

def totalScore(input_str=""):
    score = 0
    for i in best_terms:
        reg = r""+i
        add = 0
        if search(reg, input_str, re.IGNORECASE):
            if not is_all_caps(i):
                add = 10 * len(findall(reg, input_str, re.IGNORECASE))
            else:
                add = 10 * len(findall(reg, input_str))
            score += add
    for i in good_terms:
        reg = r""+i
        add = 0
        if search(reg, input_str, re.IGNORECASE):
            if not is_all_caps(i):
                add = 3 * len(findall(reg, input_str, re.IGNORECASE))
            else:
                add = 3 * len(findall(reg, input_str))
            score += add

    for i in margin_good_terms:
        reg = r""+i
        add = 0
        if search(reg, input_str, re.IGNORECASE):
            if not is_all_caps(i):
                add = len(findall(reg, input_str, re.IGNORECASE))
            else:
                add = len(findall(reg, input_str))
            score += add
    
    for i in margin_bad_terms:
        reg = r""+i
        sub = 0
        if search(reg, input_str, re.IGNORECASE):
            if not is_all_caps(i):
                sub = len(findall(reg, input_str, re.IGNORECASE))
            else:
                sub = len(findall(reg, input_str))
            score -= sub

    for i in bad_terms:
        reg = r""+i
        sub = 0
        if search(reg, input_str, re.IGNORECASE):
            if not is_all_caps(i):
                sub = 3 * len(findall(reg, input_str, re.IGNORECASE))
            else:
                sub = 3 * len(findall(reg, input_str))
            score -= sub
    
    for i in unpromising_terms:
        reg = r""+i
        sub = 0
        if search(reg, input_str, re.IGNORECASE):
            if not is_all_caps(i):
                sub = 10 * len(findall(reg, input_str, re.IGNORECASE))
            else:
                sub = 10 * len(findall(reg, input_str))
            score -= sub
    
    return score

uniques = {}

for r in range(1, no_element_sheet.nrows):
    if no_element_sheet.cell_value(r, 0) != "Patent":
        title = str(no_element_sheet.cell_value(r, 5))
        print(title)
        abstract = str(no_element_sheet.cell_value(r, 2))
        keywords = str(no_element_sheet.cell_value(r, 6))
        if title.lower() not in uniques and (containsElement(title) or containsElement(abstract)):
            '''and (totalScore(title) + totalScore(abstract) +totalScore(keywords) > 30)'''
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
            content['score'] = totalScore(title) + totalScore(abstract) +totalScore(keywords)
            tempo = {'title':allTerms(best_terms, title), 'abstract':allTerms(best_terms, abstract), 'keywords':allTerms(best_terms, keywords)}
            content['+10'] = tempo
            tempo = {'title':allTerms(good_terms, title), 'abstract':allTerms(good_terms, abstract), 'keywords':allTerms(good_terms, keywords)}
            content['+3'] = tempo
            tempo = {'title':allTerms(margin_good_terms, title), 'abstract':allTerms(margin_good_terms, abstract), 'keywords':allTerms(margin_good_terms, keywords)}
            content['+1'] = tempo
            tempo = {'title':allTerms(neutral_terms, title), 'abstract':allTerms(neutral_terms, abstract), 'keywords':allTerms(neutral_terms, keywords)}
            content['+0'] = tempo
            tempo = {'title':allTerms(margin_bad_terms, title), 'abstract':allTerms(margin_bad_terms, abstract), 'keywords':allTerms(margin_bad_terms, keywords)}
            content['-1'] = tempo
            tempo = {'title':allTerms(bad_terms, title), 'abstract':allTerms(bad_terms, abstract), 'keywords':allTerms(bad_terms, keywords)}
            content['-3'] = tempo
            tempo = {'title':allTerms(unpromising_terms, title), 'abstract':allTerms(unpromising_terms, abstract), 'keywords':allTerms(unpromising_terms, keywords)}
            content['-10'] = tempo
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
fil_sheet.write(0, 10, "Score")
fil_sheet.write(0, 11, "+10")
fil_sheet.write(0, 12, "+3")
fil_sheet.write(0, 13, "+1")
fil_sheet.write(0, 14, "+0")
fil_sheet.write(0, 15, "-1")
fil_sheet.write(0, 16, "-3")
fil_sheet.write(0, 17, "-10")
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
    fil_sheet.write(row, 10, uniques[i]['score'])
    fil_sheet.write(row, 11, str(uniques[i]['+10']))
    fil_sheet.write(row, 12, str(uniques[i]['+3']))
    fil_sheet.write(row, 13, str(uniques[i]['+1']))
    fil_sheet.write(row, 14, str(uniques[i]['+0']))
    fil_sheet.write(row, 15, str(uniques[i]['-1']))
    fil_sheet.write(row, 16, str(uniques[i]['-3']))
    fil_sheet.write(row, 17, str(uniques[i]['-10']))
    row += 1
filtered.close()