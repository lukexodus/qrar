import bs4
import requests
import pprint

res = requests.get('https://liamporritt.com/blog/100-inspirational-study-quotes')
res.raise_for_status()
soup = bs4.BeautifulSoup(res.text, 'lxml')
blockquotes = soup.select('blockquote')
rawQuotes = []
for blockquote in blockquotes:
    rawQuotes.append(blockquote.getText())

with open('quotes.py', 'w') as quotesFile:
    quotesFile.write(pprint.pformat(rawQuotes))