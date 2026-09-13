import requests
from bs4 import BeautifulSoup

def fetch_quotes_data():
    """Quotes Website Data List Function"""
    url = "https://quotes.toscrape.com"
    response = requests.get(url)
    soup = BeautifulSoup(response.text, "html.parser")
    quotes = soup.find_all("div", class_="quote")

    extracted_data = []
    for i, item in enumerate(quotes, start=1):
        text = item.find("span", class_="text").text.strip().strip('"')
        author = item.find("small", class_="author").text.strip()
        tags =", ".join([t.text.strip() for t in item.find_all("a", class_="tag")])
        extracted_data.append([i, text, author, tags])

    return extracted_data    