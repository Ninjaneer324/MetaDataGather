import requests
import xlsxwriter
apiKey = "bbcd5fe7831eb12082993dcbaaa6d72c"
url = "https://api.elsevier.com/content/search/scopus"
query = ""
#default unavailable to have all access; need to email; limit in count number varies
periodic_table = {"aluminum":"Al"}
for i in periodic_table:
    query = "(" + i + " OR " + periodic_table[i] + ") AND (precipitat* OR age)"
    print(query)
    response = requests.get(url, params={"apiKey":apiKey, "query":query, "count":25})
    g = response.json()
    print(g)