import requests
import json

url = "https://api-na.hosted.exlibrisgroup.com/primo/v1/search"
apikey = "l8xx76578ee844b74663a3cbe3b6bb3dd316"

response = requests.get(url, params={"vid":"API_GUEST_INST:API_GUEST_INST", "tab":"LibraryCatalog","scope":"MyInstitution",
"q":"creator,contains,elia zafrani,AND;any,contains,cloud computing", "qInclude":"facet_tlevel,include,online_resources|,|facet_rtype,exact,books", 
"lang":"en", "inst":"API_GUEST_INST", "apikey":apikey})
print(response.json())