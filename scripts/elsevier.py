import requests
apiKey = "bbcd5fe7831eb12082993dcbaaa6d72c"
url = "https://api.elsevier.com/content/search/scopus"
query = ""
#default unavailable to have all access; need to email; limit in count number varies
periodic_table = {"aluminum":"Al"}
for i in periodic_table:
    query = "(" + i + " OR " + periodic_table[i] + ") AND (precipitat* OR age)"
    response = requests.get(url, params={"httpAccept":"application/json","apiKey":apiKey, "query":query, "count":25})
    results = response.json()['search-results']['entry']
    abstracts = {}
    for i in results:
        colon = i['dc:identifier'].find(':')
        scopus_id = i['dc:identifier'][colon + 1:]
        scopus_abstract = "https://api.elsevier.com/content/abstract/scopus_id/" + scopus_id
        res_l = requests.get(scopus_abstract, params={"httpAccept":"application/json","apiKey":apiKey,"view":"META"})
        result_l = res_l.json()
        print(result_l)
        break