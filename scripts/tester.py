from re import search
import requests

url = "https://api.elsevier.com/content/ev/records?docId=cpx_78d813dd16ed7ac51f0M644610178163211"

apiKey = "bbcd5fe7831eb12082993dcbaaa6d72c"
#or is it the insttoken?
#API Documentation for Engineering Village search: https://dev.elsevier.com/documentation/EngineeringVillageAPI.wadl
inst_token = "4f3d2a4d46c51cbb68e83cf0b7150f45"
#endpoint for search

h = {"Accept":"application/json","X-ELS-APIKey":apiKey, "X-ELS-Insttoken":inst_token}

response = requests.get(url, headers=h)

print(response.json())