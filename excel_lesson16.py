import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import os

wb = Workbook()
ws = wb.active
ws.title = "Quotes Data"

ws.append(["No", "Quote Text", "Author"])

url = "https://quotes.toscrape.com/"

try:
    print(" website from Data start Down...")
    response = requests.get(url, timeout=10)

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")
        quotes = soup.find_all("div", class_="quote")

        for index, item in enumerate(quotes, 1):
            text = item.find("span", class_="text").text
            author = item.find("small", class_="author").text
            ws.append([index, text, author])

        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_file = os.path.join(script_dir, "scraped_quotes.xlsx")

        wb.save(output_file)
        print(f" Successfull, Data '{output_file}' in to the save")

    else:
        print(f" No success join website. Status Code: {response.status_code}")

except Exception as e:
    print(f" Falseting: {e}")                