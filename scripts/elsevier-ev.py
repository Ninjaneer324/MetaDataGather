import requests
import xlsxwriter
import xlrd
import time
from urllib.parse import unquote
#api key for authentication
apiKey = "bbcd5fe7831eb12082993dcbaaa6d72c"
#or is it the insttoken?
#API Documentation for Engineering Village search: https://dev.elsevier.com/documentation/EngineeringVillageAPI.wadl
inst_token = "4f3d2a4d46c51cbb68e83cf0b7150f45"
#endpoint for search
headers = {"Accept":"application/json","X-ELS-APIKey":apiKey, "X-ELS-Insttoken":inst_token}
url = "https://api.elsevier.com/content/ev/results"
#The list of elements/alloys we intend to query


periodic_table = {}
periodic_wb = xlrd.open_workbook("Periodic-Table.xlsx")
sheet = periodic_wb.sheet_by_index(0)
for i in range(1, 96):
    contents = {}
    contents['name'] = sheet.cell_value(i, 1)
    contents['alloys'] = sheet.cell_value(i,3)
    periodic_table[sheet.cell_value(i, 2)] = contents
#will this handle alloys with different names such as Nichrome or steel?

print("Excel Book Opening and adding work sheet...\n")
#open exccel workbook
excel_workbook = xlsxwriter.Workbook('EngineeringVillage.xlsx')
#add worksheet to workbook
worksheet = excel_workbook.add_worksheet()
#First 2 rows will detail what query format I applied
print("Writing column headers and query format...\n")
worksheet.write(0,0,"Query Format")
worksheet.write(0, 1, "(base_element OR symbol) AND (alloy_element OR symbol) AND (age* OR aging OR precipitat*) AND (phase* OR hardness OR hardening OR tensile OR microsc* or SEM OR TEM OR diffract* OR dilatom* OR (mech* AND (prop* OR response)))")
worksheet.write(2,1,"ID")
worksheet.write(2,2,"Title")
worksheet.write(2,3,"Author")
worksheet.write(2,4, "Date")

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
        query = "("+base_element+" OR "+base_symbol+") AND ("+alloy_element+" OR "+alloy_symbol+") AND (age* OR aging OR precipitat*) AND (phase* OR hardness OR hardening OR tensile OR microsc* or SEM OR TEM OR diffract* OR dilatom* OR (mech* AND (prop* OR response)))"
        #query = "(("+periodic_table[a]['name']+" OR "+a+") AND (precipitat* AND "+"(age* OR transform* OR microscop*)))"
        worksheet.write(row, 0, elem+"-"+a)
        print(query)
        #requests for search results
        response = requests.get(url, headers=headers,params={"query":query,"pageSize":100,"database":"c"}) #engineering village doesn't have count
        if response.status_code != 200:
            print("Error: HTTP", response.status_code)
            print("Closing Workbook...")
            excel_workbook.close()
            exit()
        print(response.status_code)
        results = response.json()
        worksheet.write(row + 1, 0, "Total: " + str(results['PAGE']['RESULTS-COUNT']))
        for item in results['PAGE']['PAGE-RESULTS']['PAGE-ENTRY']:
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
#close workbook'''
        row += 1

print("Closing Workbook...")
excel_workbook.close()