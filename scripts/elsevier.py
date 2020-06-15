import requests
apiKey = "bbcd5fe7831eb12082993dcbaaa6d72c"
url = "https://api.elsevier.com/content/search/sciencedirect"
query = ""
#default unavailable to have all access; need to email; limit in count number varies
periodic_table = {"aluminum":"Al"}
for i in periodic_table:
    query = "(" + i + " OR " + periodic_table[i] + ") AND (precipitat* OR age)"
    response = requests.get(url, params={"httpAccept":"application/json","apiKey":apiKey, "query":query, "count":100})
    results = response.json()['search-results']['entry']
    