 
import requests
import xlsxwriter
import xlrd
import time
#api key for authentication
apiKey = "bbcd5fe7831eb12082993dcbaaa6d72c"
#or is it the insttoken?
#API Documentation for Engineering Village search: https://dev.elsevier.com/documentation/EngineeringVillageAPI.wadl
access_token = "4f3d2a4d46c51cbb68e83cf0b7150f45"
#endpoint for search
headers = {"Accept":"application/json","X-ELS-APIKey":apiKey, "X-ELS-Insttoken":access_token}
url = "https://api.elsevier.com/content/search/sciencedirect"
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
excel_workbook = xlsxwriter.Workbook('ScienceDirect.xlsx')
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
        #query = "("+base_element+" OR "+base_symbol+") AND ("+alloy_element+" OR "+alloy_symbol+") AND (age* OR aging OR precipitat*) AND (phase* OR hardness OR hardening OR tensile OR microsc* or SEM OR TEM OR diffract* OR dilatom* OR (mech* AND (prop* OR response)))"
        query = "(("+periodic_table[a]['name']+" OR "+a+") AND (precipitat* AND "+"(age* OR transform* OR microscop*)))"
        worksheet.write(row, 0, elem+"-"+a)
        print(query)
        #requests for search results
        response = requests.get(url, headers=headers,params={"query":query,"count":200}) #engineering village doesn't have count
        
        if response.status_code != 200:
            print("Error: HTTP", response.status_code)
            print("Closing Workbook...")
            excel_workbook.close()
        results = response.json()
        print(response.status_code)
        #writes what element is currently being queried into worksheet
        worksheet.write(row + 1, 0, "Total: " + str(results['search-results']['opensearch:totalResults']))
        for r in results['search-results']['entry']:
            #checks that said meta data exists and writes it to the excel spread sheet cell if it does and 
            #to the missing information row if it doesn't
            if 'dc:identifier' in r:
                id = r['dc:identifier']
                worksheet.write(row, 1, id)
            else:
                worksheet.write(row,1,"Missing")
            if 'dc:title' in r:
                title = r['dc:title']
                worksheet.write(row, 2, title)
            else:
                worksheet.write(row, 2, "Missing")
            
            if 'dc:creator' in r:
                creator = r['dc:creator']
                worksheet.write(row, 3, creator)
            else:
                missing.write(row, 3, "Missing")
                
            if 'prism:coverDate' in r:
                coverDate = r['prism:coverDate']
                worksheet.write(row, 4, coverDate)
            else:
                worksheet.write(row,4,"Missing")
                
            if 'load-date' in r:
                load_date = r['load-date']
                worksheet.write(row, 5, load_date)
            else:
                worksheet.write(row,5,"Missing")
            row += 1
        time.sleep(2)
        row += 1
print("Closing Workbook...")
excel_workbook.close()