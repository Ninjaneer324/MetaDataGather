import xlsxwriter
import xlrd
import requests
import time
from dateutil.parser import parse
from math import ceil
import re

apiKey = "bbcd5fe7831eb12082993dcbaaa6d72c"
#or is it the insttoken?
#API Documentation for Engineering Village search: https://dev.elsevier.com/documentation/EngineeringVillageAPI.wadl
inst_token = "4f3d2a4d46c51cbb68e83cf0b7150f45"
#endpoint for search
headers = {"Accept":"application/json","X-ELS-APIKey":apiKey, "X-ELS-Insttoken":inst_token}
url = "https://api.elsevier.com/content/ev/results"

periodic_table = {}
periodic_wb = xlrd.open_workbook("Periodic-Table.xlsx")
sheet = periodic_wb.sheet_by_index(0)
for i in range(1, 96):
    contents = {}
    contents['name'] = sheet.cell_value(i, 1)
    contents['alloys'] = sheet.cell_value(i,3)
    contents['pos'] = i
    periodic_table[sheet.cell_value(i, 2)] = contents

'''nxn_table = xlsxwriter.Workbook('FinalProduct.xlsx')
worksheet = nxn_table.add_worksheet()
sheet = periodic_wb.sheet_by_index(0)
for i in range(1, 96):
    stuff = sheet.cell_value(i, 2)
    worksheet.write(0, i, stuff)
    worksheet.write(i, 0, stuff)'''
del periodic_wb

def searchBaseAlloy(base_n = "", base_s = "", alloy_n = "", alloy_s = "", start=0):
    name_file = base_s+"-"+alloy_s+".xlsx"
    if base_n == "" and base_s == "" and alloy_n == "" and alloy_s == "":
        name_file = "NoElement.xlsx"
    workbook = xlsxwriter.Workbook(name_file)
    worksheet = workbook.add_worksheet()
    row = 0
    worksheet.write(row, 0, "Query Format")
    f_mat = "\"{0} alloys\" AND \"-*{3}\" AND (age* OR aging OR precipitat* OR inclusion* OR dispersoid* OR \"solid solution\" OR solub* OR solutionize OR new*phase) AND (hardness OR hardening OR harden* OR strength*) AND (phase* OR tensile OR microsc* or SEM OR TEM OR diffract* OR dilatom* OR solvus OR (mech* AND (prop* OR response))) NOT (biol* OR diseas* OR cancer OR aqueous* OR ceramic OR \" Fe-*\" OR steel OR \" Al-*\" OR \" Mg-*\" OR \"0-9+\" OR \"IV*\" OR \"VI*\")"
    if base_n == alloy_n and base_n != "":
        f_mat = "\"pure "+base_n+"\" AND (age* OR aging OR precipitat* OR inclusion* OR dispersoid* OR \"solid solution\" OR solub* OR solutionize OR new*phase) AND (hardness OR hardening OR harden* OR strength*) AND (phase* OR tensile OR microsc* or SEM OR TEM OR diffract* OR dilatom* OR solvus OR (mech* AND (prop* OR response))) NOT (biol* OR diseas* OR cancer OR aqueous* OR ceramic OR \" Fe-*\" OR steel OR \" Al-*\" OR \" Mg-*\" OR \"0-9+\" OR \"IV*\" OR \"VI*\")"
    elif base_n == "" and base_s == "" and alloy_n == "" and alloy_s == "":
        f_mat = "(age* OR aging OR precipitat* OR inclusion* OR dispersoid* OR \"solid solution\" OR solub* OR solutionize OR new*phase) AND (hardness OR hardening OR harden* OR strength*) AND (phase* OR tensile OR microsc* or SEM OR TEM OR diffract* OR dilatom* OR solvus OR (mech* AND (prop* OR response))) NOT (biol* OR diseas* OR cancer OR aqueous* OR ceramic OR \" Fe-*\" OR steel OR \" Al-*\" OR \" Mg-*\" OR \"0-9+\" OR \"IV*\" OR \"VI*\")"
    worksheet.write(row, 1, f_mat)
    row = 2
    query = f_mat.format(base_n, base_s, alloy_n, alloy_s)
    print(query)
    response = requests.get(url, headers=headers,params={"query":query, "offset":start, "pageSize":100, "database":"c","sortField":"relevance"})
    print("Page", start + 1,response.status_code)
    if response.status_code != 200:
        print("Error: HTTP", response.status_code)
        print("Closing Workbook...")
        temp_workbook.close()
        exit()

    #first page of results
    worksheet.write(row, 0, base_s+"-"+alloy_s)
    exclude = 0
    results = response.json()
    total_results = results['PAGE']['RESULTS-COUNT']
    if 'PAGE-RESULTS' in results['PAGE']:
        for item in results['PAGE']['PAGE-RESULTS']['PAGE-ENTRY']:
            t = item['EI-DOCUMENT']['DOCUMENTPROPERTIES']['TI']
            if "<inf>" in t:
                exclude += 1
                continue
            if 'EI-DOCUMENT' in item and 'DOCUMENTPROPERTIES' in item['EI-DOCUMENT']:
                ids = {}
                if 'DO' in item['EI-DOCUMENT']['DOCUMENTPROPERTIES']:
                    id = item['EI-DOCUMENT']['DOCUMENTPROPERTIES']['DO']
                    ids['DOI'] = id
                    worksheet.write(row, 1, str(ids))
                else:
                    worksheet.write(row, 1, "missing")
                        
                if 'TI' in item['EI-DOCUMENT']['DOCUMENTPROPERTIES']:
                    title = item['EI-DOCUMENT']['DOCUMENTPROPERTIES']['TI']
                    worksheet.write(row, 2, title)
                else:
                    worksheet.write(row, 2, "missing")
                        
                if 'SD' in item['EI-DOCUMENT']['DOCUMENTPROPERTIES']:
                    date = item['EI-DOCUMENT']['DOCUMENTPROPERTIES']['SD']
                    worksheet.write(row, 4, date)
                else:
                    worksheet.write(row,4,"missing")
            else:
                    worksheet.write(row, 1, "missing")
                    worksheet.write(row, 2, "missing")
                    worksheet.write(row,4,"missing")

            if 'AUS' in item['EI-DOCUMENT'] and 'AU' in item['EI-DOCUMENT']['AUS']:
                s = ""
                for i in range(len(item['EI-DOCUMENT']['AUS']['AU'])):
                    s += item['EI-DOCUMENT']['AUS']['AU'][i]['NAME']
                    if i != len(item['EI-DOCUMENT']['AUS']['AU']) - 1:
                        s += ";"
                worksheet.write(row, 3, s)
            else:
                worksheet.write(row, 3, "missing")
                
            if 'DOI' not in ids and 'DOCUMENTOBJECTS' in item['EI-DOCUMENT'] and 'CITEDBY' in item['EI-DOCUMENT']['DOCUMENTOBJECTS'] and 'DOI' in item['EI-DOCUMENT']['DOCUMENTOBJECTS']['CITEDBY']:
                id = unquote(item['EI-DOCUMENT']['DOCUMENTOBJECTS']['CITEDBY']['DOI'])
                ids['DOI'] = id
                worksheet.write(row, 1, str(ids))
            elif 'DOC' in item['EI-DOCUMENT'] and 'DOC-ID' in item['EI-DOCUMENT']['DOC']:
                id = item['EI-DOCUMENT']['DOC']['DOC-ID']
                ids['DOC-ID'] = id
                worksheet.write(row, 1, str(ids))
            row += 1
        time.sleep(2)

        pages = ceil(total_results / 100)
        for i in range(start + 1, pages):
            response = requests.get(url, headers=headers,params={"query":query,"offset":i,"pageSize":100,"database":"c","sortField":"relevance"}) #engineering village doesn't have count
            print("Page",i + 1,response.status_code)
            if response.status_code != 200:
                print("Error: HTTP", response.status_code)
                print("Closing Workbook...")
                excel_workbook.close()
                doc_id_file.close()
                exit()

            results = response.json()
            for item in results['PAGE']['PAGE-RESULTS']['PAGE-ENTRY']:
                t = item['EI-DOCUMENT']['DOCUMENTPROPERTIES']['TI']
                if "<inf>" in t:
                    exclude += 1
                    continue
                if 'EI-DOCUMENT' in item and 'DOCUMENTPROPERTIES' in item['EI-DOCUMENT']:
                    ids = {}
                    if 'DO' in item['EI-DOCUMENT']['DOCUMENTPROPERTIES']:
                        id = item['EI-DOCUMENT']['DOCUMENTPROPERTIES']['DO']
                        ids['DOI'] = id
                        worksheet.write(row, 1, str(ids))
                    else:
                        worksheet.write(row, 1, "missing")
                            
                    if 'TI' in item['EI-DOCUMENT']['DOCUMENTPROPERTIES']:
                        title = item['EI-DOCUMENT']['DOCUMENTPROPERTIES']['TI']
                        worksheet.write(row, 2, title)
                    else:
                        worksheet.write(row, 2, "missing")
                            
                    if 'SD' in item['EI-DOCUMENT']['DOCUMENTPROPERTIES']:
                        date = item['EI-DOCUMENT']['DOCUMENTPROPERTIES']['SD']
                        worksheet.write(row, 4, date)
                    else:
                        worksheet.write(row,4,"missing")
                else:
                    worksheet.write(row, 1, "missing")
                    worksheet.write(row, 2, "missing")
                    worksheet.write(row,4,"missing")

                if 'AUS' in item['EI-DOCUMENT'] and 'AU' in item['EI-DOCUMENT']['AUS']:
                    s = ""
                    for i in range(len(item['EI-DOCUMENT']['AUS']['AU'])):
                        s += item['EI-DOCUMENT']['AUS']['AU'][i]['NAME']
                        if i != len(item['EI-DOCUMENT']['AUS']['AU']) - 1:
                            s += ";"
                    worksheet.write(row, 3, s)
                else:
                    worksheet.write(row, 3, "missing")
                    
                if 'DOI' not in ids and 'DOCUMENTOBJECTS' in item['EI-DOCUMENT'] and 'CITEDBY' in item['EI-DOCUMENT']['DOCUMENTOBJECTS'] and 'DOI' in item['EI-DOCUMENT']['DOCUMENTOBJECTS']['CITEDBY']:
                    id = unquote(item['EI-DOCUMENT']['DOCUMENTOBJECTS']['CITEDBY']['DOI'])
                    ids['DOI'] = id
                    worksheet.write(row, 1, str(ids))
                elif 'DOC' in item['EI-DOCUMENT'] and 'DOC-ID' in item['EI-DOCUMENT']['DOC']:
                    id = item['EI-DOCUMENT']['DOC']['DOC-ID']
                    ids['DOC-ID'] = id
                    worksheet.write(row, 1, str(ids))
                row += 1
            time.sleep(2)
        total_results -= exclude
        if (i + 1) % 100 == 0:
            time.sleep(10)
    else:
        time.sleep(2)
    worksheet.write(3, 0, total_results)
    workbook.close()
    return total_results

searchBaseAlloy()
#workbook_array = [["" for i in range(1, 96)] for j in range(1, 96)]

'''for i in periodic_table:
    for j in periodic_table:
        cell_str = ""
        total = searchBaseAlloy(periodic_table[i]['name'], i, periodic_table[j]['name'], j)
        filename = i+"-"+j+".xlsx"
        if total != 0:
            read_workbook = xlrd.open_workbook(filename)
            sheet = read_workbook.sheet_by_index(0)
            len_cnt = 0
            for row in range(2, sheet.nrows):
                date = sheet.cell_value(row, 4)
                year = None
                years = re.findall(r"[0-9]{4,7}(?![0-9])", date)
                if len(years) != 0:
                    year = years[0] #figure this out 
                authorlist = sheet.cell_value(row, 3)
                comma = authorlist.find(',')
                author = authorlist[0:comma]
                cell_str += str(year) + author
                len_cnt += 1
                if row != sheet.nrows - 1 and len_cnt < 10:
                    cell_str += ";"
                else:
                    break
        workbook_array[periodic_table["Mo"]["pos"] - 1][periodic_table[i]["pos"] - 1] = cell_str'''

#nxn_table.close()