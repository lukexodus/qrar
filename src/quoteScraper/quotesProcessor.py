import quotes
import pprint

strippedQuotes = []
for quote in quotes.quotes:
    strippedQuotes.append(quote[2:-2])

with open('quotes.py', 'a') as quotesFile:
    quotesFile.write(pprint.pformat(strippedQuotes))