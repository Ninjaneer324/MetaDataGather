'''import gensim
from gensim import corpora
from pprint import pprint
from gensim.utils import simple_preprocess
from smart_open import smart_open
import os

# Create gensim dictionary form a single tet file
dictionary = corpora.Dictionary(simple_preprocess(line, deacc=True) for line in open('sample.txt', encoding='utf-8'))

# Token to Id map
dictionary.token2id
print(dictionary.token2id)
#> {'according': 35,
#>  'and': 22,
#>  'appointment': 23,
#>  'army': 0,
#>  'as': 43,
#>  'at': 24,
#>   ...
#> }'''
import requests
apiKey = "bbcd5fe7831eb12082993dcbaaa6d72c"

inst_token = "4f3d2a4d46c51cbb68e83cf0b7150f45"

url = "https://api.elsevier.com/content/ev/records"
headers = {"Accept":"application/json","X-ELS-APIKey":apiKey, "X-ELS-Insttoken":inst_token}
r = requests.get(url, headers=headers,params={"docId":"cpx_6e3d601232e2239f5M5bcd2061377553"})
results = r.json()
print(results['PAGE']['PAGE-RESULTS']['PAGE-ENTRY'][0]['EI-DOCUMENT']['DOCUMENTPROPERTIES']['AB'])