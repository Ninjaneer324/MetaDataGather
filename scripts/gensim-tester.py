import gensim
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
#> }
