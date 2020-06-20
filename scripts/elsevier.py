import requests
import xlsxwriter
#api key for authentication
apiKey = "bbcd5fe7831eb12082993dcbaaa6d72c"
#endpoint for search
url = "https://api.elsevier.com/content/search/sciencedirect"
query = ""
#The list of elements/alloys we intend to query
periodic_table = {"aluminum":"Al","iron":"Fe"}
#will this handle alloys with different names such as Nichrome or steel?
#open exccel workbook
excel_workbook = xlsxwriter.Workbook('Elsevier.xlsx')
#add worksheet to workbook
worksheet = excel_workbook.add_worksheet()
#First 2 rows will detail what query format I applied
worksheet.write(0,0,"Query Format")
worksheet.write(0, 1, "(<element> OR <element_symbol>) AND ((precipitat* OR age) harden*)")

#query portion
row = 2
#will hold metadata that will say what information that might be missing from each query
for i in periodic_table:
    #formats query for each element in the "periodic table"
    query = "(" + i + " OR " + periodic_table[i] + ") AND ((precipitat* OR age) harden*)"
    #requests for search results
    response = requests.get(url, params={"httpAccept":"application/json","apiKey":apiKey, "query":query, "count":100})
    results = response.json()['search-results']['entry']
    #writes what element is currently being queried into worksheet
    worksheet.write(row, 0, periodic_table[i])
    worksheet.write(row + 1, 0, "Missing Information")
    col = 1
    for i in results:
        cell = {}
        missing = []
        #checks that said meta data exists and writes it to the excel spread sheet cell if it does and 
        #to the missing information row if it doesn't
        if 'dc:identifier' in i:
            id = i['dc:identifier']
            cell['id'] = id
        else:
            missing.append('id')

        if 'dc:title' in i :
            title = i['dc:title']
            cell['title'] = title
        else:
            missing.append('title')
    
        if 'dc:creator' in i:
            creator = i['dc:creator']
            cell['author'] = creator
        else:
            missing.append('author')
        
        if 'prism:coverDate' in i:
            cell['prism:coverDate'] = i['prism:coverDate']
        else:
            missing.append('prism:coverDate')
        
        if 'load-date' in i:
            cell['load-date'] = i['load-date']
        else:
            missing.append('load-date')
        worksheet.write(row, col, str(cell))
        worksheet.write(row + 1, col, str(missing))
        col += 1

    row+=3
#close workbook
excel_workbook.close()