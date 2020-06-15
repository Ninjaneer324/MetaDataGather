from bs4 import BeautifulSoup #parsing of web pages
import datetime
from tinydb import TinyDB, Query #data storage stuff
import urllib3 #makes http requests
import xlsxwriter #makes excel sheet

# note for me to know it started
print('hloe')

# ignores any SSL certificate warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# the url to be searched
url = 'https://link.springer.com/article/10.1007/s11661-002-0124-3'

# gets the source info
def make_soup(url):
    http = urllib3.PoolManager()
    r = http.request('GET', url, retries=urllib3.util.Retry(total=13, redirect=13))
    print(r.data)
    return BeautifulSoup(r.data,'lxml')
soup = make_soup(url)

# pulls what we want from the soup
title1 = soup.find("meta", attrs={'name':"dc.title"}).get("content")
journal = soup.find("meta", attrs={'name':"dc.source"}).get("content")
publisher = soup.find("meta", attrs={'name':"dc.publisher"}).get("content")
abstract = soup.find("meta", attrs={'name':"dc.description"}).get("content")
fauthor = soup.find("meta", attrs={'name':"citation_author"}).get("content")
doi = soup.find("meta", attrs={'name':"DOI"}).get("content")
url = soup.find("meta", attrs={'name':"citation_fulltext_html_url"}).get("content")
date = soup.find("meta", attrs={'name':"citation_publication_date"}).get("content")

# make spreadsheet
Headlines = ["title", "fauthor", "date", "journal", "doi", "url", "publisher", "abstract"]
row = 0
workbook = xlsxwriter.Workbook('single_citation_scrape.xlsx')
worksheet = workbook.add_worksheet()
# add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

for col, title in enumerate(Headlines):
    worksheet.write(row, col, title, bold)
col = 0
row = 1
content = [title1, fauthor, date, journal, doi, url, publisher, abstract]
for col in range(len(Headlines)):
    worksheet.write(row, col, content[col])
    col += 1

workbook.close()