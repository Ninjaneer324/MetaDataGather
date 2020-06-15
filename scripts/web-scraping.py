import requests
from bs4 import BeautifulSoup #parsing of web page


url = "https://scholar.google.com/scholar?start=0&q=aluminum+precipitation&hl=en&as_sdt=0,44"
headers = {'User-Agent':"Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:76.0) Gecko/20100101 Firefox/76.0"}
pages_and_content = {}
change = url.find("start=") + len("start=")
#now try to loop through pages
#Error HTTP 429 occurs to often; web scraping is probably just not enough
for k in range(10):
    p = []
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        print("Error Status Code: " + str(response.status_code))
        exit()
    soup = BeautifulSoup(response.text, "html.parser")
    r = soup.findAll('div', attrs={"class":"gs_r gs_or gs_scl"})
    for i in r:
        c = {}
        if i.find('h3').find('a') is not None:
            c['title'] = i.find('h3').find('a').text
            if i.find('h3').find('span', attrs={'class':'gs_ctc'}) is not None:
                c['type'] = i.find('h3').find('span', attrs={'class':'gs_ctc'}).text
            else:
                c['type'] = 'link'
        elif i.find('h3').find('span') is not None:
            getAll = i.find('h3').findAll('span')
            c['title'] = getAll[1].text
            c['type'] = getAll[0].text
        p.append(c)
    pages_and_content["Page " + str(k + 1)] = p
    first_part = url[0:change]
    next_10 = int(url[change:url.find("&")]) + 10
    last_part = url[change + len(str(next_10 - 10)):]
    url = first_part + str(next_10) + last_part