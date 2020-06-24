import requests
import xlsxwriter
import xlrd
#api key for authentication
apiKey = "bbcd5fe7831eb12082993dcbaaa6d72c"
#or is it the insttoken?
#API Documentation for Engineering Village search: https://dev.elsevier.com/documentation/EngineeringVillageAPI.wadl
access_token = ""
#endpoint for search
#url = "https://api.elsevier.com/content/search/sciencedirect"
url = "https://api.elsevier.com/content/ev/results"
query = ""
#The list of elements/alloys we intend to query


periodic_table = {}
periodic_wb = xlrd.open_workbook("/home/ninjaneer324/Documents/MetaDataGather/Period LIST.xlsx")
sheet = periodic_wb.sheet_by_index(0)
for i in range(1, 96):
    contents = {}
    contents['name'] = sheet.cell_value(i, 1)
    contents['alloys'] = sheet.cell_value(i,3)
    periodic_table[sheet.cell_value(i, 2)] = contents
#will this handle alloys with different names such as Nichrome or steel?


#open exccel workbook
excel_workbook = xlsxwriter.Workbook('Elsevier.xlsx')
#add worksheet to workbook
worksheet = excel_workbook.add_worksheet()
#First 2 rows will detail what query format I applied
worksheet.write(0,0,"Query Format")
worksheet.write(0, 1, "(((<base_element> OR <symbol>) AND (<alloy> OR <symbol>)) AND (precipitat* AND "+
"(age* OR transform* OR microscop*))) NOT (aqueous OR bio* OR disease*)")


#query portion
row = 2
#will hold metadata that will say what information that might be missing from each query
for elem in periodic_table:
    #formats query for each element in the "periodic table"
    if len(periodic_table[elem]['alloys']) == 0:
        print(elem)
        print("continued\n")
        continue
    alloys = periodic_table[elem]['alloys'].split(', ')
    for a in alloys:
        query = "((("+periodic_table[elem]['name']+" OR "+elem+") AND ("+periodic_table[a]['name']+" OR "+a+")) AND (precipitat* AND "+"(age* OR transform* OR microscop*))) NOT (aqueous OR bio* OR disease*)"
        print(query)
        #requests for search results
        response = requests.get(url, params={"httpAccept":"application/json","apiKey":apiKey, "access_token":access_token,"query":query})
        results = response.json()['search-results']['entry']
        #writes what element is currently being queried into worksheet
        worksheet.write(row, 0, elem + "-" +a)
        worksheet.write(row + 1, 0, "Missing Information")
        col = 1
        for r in results:
            cell = {}
            missing = []
            #checks that said meta data exists and writes it to the excel spread sheet cell if it does and 
            #to the missing information row if it doesn't
            if 'dc:identifier' in r:
                id = r['dc:identifier']
                cell['id'] = id
            else:
                missing.append('id')

            if 'dc:title' in r:
                title = r['dc:title']
                cell['title'] = title
            else:
                missing.append('title')
        
            if 'dc:creator' in r:
                creator = r['dc:creator']
                cell['author'] = creator
            else:
                missing.append('author')
            
            if 'prism:coverDate' in r:
                cell['prism:coverDate'] = r['prism:coverDate']
            else:
                missing.append('prism:coverDate')
            
            if 'load-date' in r:
                cell['load-date'] = r['load-date']
            else:
                missing.append('load-date')
            worksheet.write(row, col, str(cell))
            worksheet.write(row + 1, col, str(missing))
            col += 1
        row+=2
    row+=1
#close workbook'''
    print()
excel_workbook.close()