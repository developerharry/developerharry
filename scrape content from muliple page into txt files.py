from bs4 import BeautifulSoup
import requests
import pandas as pd

headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.12; rv:55.0) Gecko/20100101 Firefox/55.0',
}

df = pd.read_excel(r'C:\Users\user\Downloads\Input.xlsx', sheet_name=0) # can also index sheet by name or fetch all sheets
mylist = df['URL'].tolist()

i = 0
for x in range(len(mylist)):
    url = mylist[x]
    markup = requests.get(url, headers=headers).text
    soup = BeautifulSoup(markup, "html.parser")

    site_title = soup.find("h1")

    i = i+1
    file = open("{}.txt".format(i), "w", encoding="utf-8")
    file.write(site_title.text)

    site_content = soup.find('div', class_= 'td-post-content')
    for content in site_content:
        file.write(content.text)

    file.close()

