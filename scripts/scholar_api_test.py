import requests, json


def inRange(input, index):
    return -1 < input < len(index)

def getDate(input):
    for c in range(len(input) - 1, -1, -1):
        if (inRange(input, i) for i in range(c, c+4)) and input[c:c+4].isdigit():
            return input[c:c+4]
    return None

results_dict = []

#need to pay for serpapi after the 15 day trial period, $50/month
api_key = "62cf3415aa6986ce1d4338784dea442e12d5962b7b58a0524d888aed8e866c8c"
url = "https://serpapi.com/search?engine=google_scholar"
getInfo = requests.get(url, params={"api_key":api_key,"q":"aluminum precipitation"})
paging = requests.get(url)
results = getInfo.json()["organic_results"]
for r in results:
    elem = {}
    elem["title"] = r["title"]
    if "link" in r:
        elem["link"] = r["link"]
    if "resources" in r:
        elem["resources"] = r["resources"]
    if "publication_info" in r:
        if "summary" in r["publication_info"]:
            date = getDate(r["publication_info"]["summary"])
            if date is not None:
                elem["date"] = date 
        if "authors" in r["publication_info"]:
            elem["author"] = r["publication_info"]["authors"][0]["name"]
    #abstract unattainable without redirecting
    #how to handle other cases like citations
    #what to put in the search query and how to handle permutations of the different elements
    #how to handle multiple languages-API actually has a parameter for this
'''
from serpapi.google_search_results import GoogleSearchResults

params = {
  "api_key": "62cf3415aa6986ce1d4338784dea442e12d5962b7b58a0524d888aed8e866c8c",
  "engine": "google_scholar",
  "hl": "en",
}

client = GoogleSearchResults(params)
results = client.get_dict()
'''