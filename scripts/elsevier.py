import requests
import xlsxwriter
apiKey = "bbcd5fe7831eb12082993dcbaaa6d72c"
url = "https://api.elsevier.com/content/search/sciencedirect"
query = ""
#default unavailable to have all access; need to email; limit in count number varies
periodic_table = {"aluminum":"Al"}
excel_workbook = xlsxwriter.Workbook('Elsevier.xlsx')
worksheet = excel_workbook.add_worksheet()
worksheet.write(0,0,"Query Format")
worksheet.write(0, 1, "(<element> OR <element_symbol>) AND ((precipitat* OR age) harden*)")
row = 1
missing_information = {}
for i in periodic_table:
    query = "(" + i + " OR " + periodic_table[i] + ") AND ((precipitat* OR age) harden*)"
    response = requests.get(url, params={"httpAccept":"application/json","apiKey":apiKey, "query":query, "count":100})
    results = response.json()['search-results']['entry']
    worksheet.write(row, 0, periodic_table[i])
    col = 1
    for i in results:
        cell = {}
        missing = []
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
        missing_information[str(col)] = missing
        worksheet.write(row, col, str(cell))
        col += 1
    row += 1
    worksheet.write(row, 0, "Missing information")
    for i in missing_information:
        if len(missing_information[i]) > 0:
            worksheet.write(row, int(i), str(missing_information[i]))
    missing_information.clear()
    row+=1
    
excel_workbook.close()

