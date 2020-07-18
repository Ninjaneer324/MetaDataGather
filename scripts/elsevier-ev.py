import requests
import xlsxwriter
import xlrd
import time
from urllib.parse import unquote
from math import ceil
#api key for authentication
apiKey = "bbcd5fe7831eb12082993dcbaaa6d72c"
#or is it the insttoken?
#API Documentation for Engineering Village search: https://dev.elsevier.com/documentation/EngineeringVillageAPI.wadl
inst_token = "4f3d2a4d46c51cbb68e83cf0b7150f45"
#endpoint for search
headers = {"Accept":"application/json","X-ELS-APIKey":apiKey, "X-ELS-Insttoken":inst_token}
url = "https://api.elsevier.com/content/ev/results"
#The list of elements/alloys we intend to query
total_searches = 0

periodic_table = {}
periodic_wb = xlrd.open_workbook("Periodic-Table.xlsx")
sheet = periodic_wb.sheet_by_index(0)
for i in range(1, 96):
    contents = {}
    contents['name'] = sheet.cell_value(i, 1)
    contents['alloys'] = sheet.cell_value(i,3)
    periodic_table[sheet.cell_value(i, 2)] = contents
#will this handle alloys with different names such as Nichrome or steel?

periodic_table['Al']['alloys'] = {}

print("Excel Book Opening and adding work sheet...\n")
#open exccel workbook
excel_workbook = xlsxwriter.Workbook('EngineeringVillage-basename-alloysymbol.xlsx')
#add worksheet to workbook
worksheet = excel_workbook.add_worksheet()
#First 2 rows will detail what query format I applied
print("Writing column headers and query format...\n")
worksheet.write(0,0,"Query Format")
#f_mat = "({0} OR {1}) AND ({2} OR {3}) AND (age* OR aging OR precipitat*) AND (phase* OR hardness OR hardening OR tensile OR microsc* or SEM OR TEM OR diffract* OR dilatom* OR (mech* AND (prop* OR response))) NOT (biol* OR diseas* OR cancer OR aqueous* OR ceramic OR \" Fe-*\" OR steel OR \" Al-*\" OR \" Mg-*\" OR \"0-9+\" OR \"IV*\" OR \"VI*\")"
#f_mat = "{0} AND {2} AND (age* OR aging OR precipitat*) AND (phase* OR hardness OR hardening OR tensile OR microsc* or SEM OR TEM OR diffract* OR dilatom* OR (mech* AND (prop* OR response))) NOT (biol* OR diseas* OR cancer OR aqueous OR ceramic OR \" Fe-*\" OR steel OR \" Al-*\" OR \" Mg-*\" OR \"0-9+\" OR \"IV*\" OR \"VI*\")"
#f_mat = "\" {1}-*\" AND \"-*{3}\" AND (age* OR aging OR precipitat*) AND (phase* OR hardness OR hardening OR tensile OR microsc* or SEM OR TEM OR diffract* OR dilatom* OR (mech* AND (prop* OR response))) NOT (biol* OR diseas* OR cancer OR aqueous* OR ceramic OR laser OR \" Fe-*\" OR steel OR \" Al-*\" OR \" Mg-*\" OR \"0-9+\" OR \"(I*)\" OR \"(IV)\" OR \"(V*)\")"
f_mat = "\"{0} alloys\" AND \"-*{3}\" AND (age* OR aging OR precipitat*) AND (phase* OR hardness OR hardening OR tensile OR microsc* or SEM OR TEM OR diffract* OR dilatom* OR (mech* AND (prop* OR response))) NOT (biol* OR diseas* OR cancer OR aqueous* OR ceramic OR \" Fe-*\" OR steel OR \" Al-*\" OR \" Mg-*\" OR \"0-9+\" OR \"IV*\" OR \"VI*\")"
worksheet.write(0, 1, f_mat.format("base_element","base_symbol", "alloy_element", "alloy_symbol"))
worksheet.write(2,1,"ID")
worksheet.write(2,2,"Title")
worksheet.write(2,3,"Author")
worksheet.write(2,4, "Date")
doc_id_file = open("doc_ids.txt","w")
print("Begin Mass Query...\n")
print("Loading...\n\n")
#query portion
row = 3
#will hold metadata that will say what information that might be missing from each query
for elem in periodic_table:
    #formats query for each element in the "periodic table"
    if len(periodic_table[elem]['alloys']) == 0:
        continue
    alloys = periodic_table[elem]['alloys'].split(', ')
    for a in alloys:
        base_element = periodic_table[elem]['name']
        base_symbol = elem
        alloy_element = periodic_table[a]['name']
        alloy_symbol = a
        query = f_mat.format(base_element, base_symbol, alloy_element, alloy_symbol)
        worksheet.write(row, 0, elem+"-"+a)
        print(query)
        #requests for search results
        response = requests.get(url, headers=headers,params={"query":query,"pageSize":100,"database":"c","sortField":"relevance"}) #engineering village doesn't have count
        print("Page 1",response.status_code)
        if response.status_code != 200:
            print("Error: HTTP", response.status_code)
            print("Closing Workbook...")
            excel_workbook.close()
            doc_id_file.close()
            exit()

        #first page of results
        results = response.json()
        total_results = results['PAGE']['RESULTS-COUNT']
        if 'PAGE-RESULTS' in results['PAGE']:
            worksheet.write(row + 1, 0, "Total: " + str(total_results))
            for item in results['PAGE']['PAGE-RESULTS']['PAGE-ENTRY']:
                t = item['EI-DOCUMENT']['DOCUMENTPROPERTIES']['TI']
                if "<inf>" in t:
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
                    doc_id_file.write(elem+"-"+a+"\t"+id+"\n")
                    worksheet.write(row, 1, str(ids))
                row += 1
            time.sleep(2)


            #next 8 pages if any
            pages = min(8, ceil(total_results / 100))
            for i in range(1, pages):
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
                        doc_id_file.write(elem+"-"+a+"\t"+id+"\n")
                        worksheet.write(row, 1, str(ids))
                    row += 1
                time.sleep(2)
#close workbook'''
        row += 1

print("Closing Workbook...")
excel_workbook.close()
doc_id_file.close()