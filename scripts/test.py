import requests
import json

url = "https://api-na.hosted.exlibrisgroup.com/primo/v1/search"
apikey = "l8xx76578ee844b74663a3cbe3b6bb3dd316"
#I can't use my own apikey for some reason???
#because my idiot head forgot to change the api area to Primo Search
response = requests.get(url, params={'vid':'API_GUEST_INST:API_GUEST_INST', 'tab':'LibraryCatalog','scope':'MyInstitution',
'q':'creator,contains,elia zafrani,AND;any,contains,cloud computing', 'lang':'en', 'inst':'API_GUEST_INST', 'apikey':apikey})

c = response.json()

print(c)