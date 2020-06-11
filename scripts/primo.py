import requests
import json
import xlsxwriter

#these are the variables to set for a guest user
#I am unsure of some of these variables are part of a normal search query or if you have
#to have your institution provide this
vid = 'API_GUEST_INST:API_GUEST_INST'
tab = 'LibraryCatalog'
scope = 'MyInstitution'
#for q, always leave it in single quotes, if you wanna look for an exact phrase, like the example below, use double quotes
q = ''
inst = 'API_GUEST_INST'
#Jessica or whoever is running the program needs to get their own apikey
apikey = 'l8xx76578ee844b74663a3cbe3b6bb3dd316'

url = "https://api-na.hosted.exlibrisgroup.com/primo/v1/search"

periodic_table = []
workbook = xlsxwriter.Workbook("output.xlsx")
worksheet = workbook.add_worksheet()
row = 0
for e in periodic_table:
    q = 'any,contains,precipitat*,AND;any,contains,' + e
    #does the hard work of adding queries for you
    response = requests.get(url, params={'vid':vid, 'tab': tab, 'scope':scope,'q':q,'limit':100,'inst':inst,'apikey':apikey})
    full_response = response.json()
    worksheet.write(row,0,e)
    if response.status_code != 200:
        print("Error: HTTP " + str(response.status_code))
    else: 
        docs = response.json()['docs']
        for i in range(len(docs)):
            result = docs[i]['pnx']['sort']['title'][0] + ","
            if 'creationdate' in docs[i]['pnx']['sort'] and 'author' in docs[i]['pnx']['sort']:
                date = docs[i]['pnx']['sort']['creationdate'][0]
                author_full = ",\"" + docs[i]['pnx']['sort']['author'][0] + "\""
                result += date + author_full
            worksheet.write(row, i + 1, result)
    row += 1
workbook.close()