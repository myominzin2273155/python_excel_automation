import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import os

wb = Workbook()
ws = wb.active
ws.title = "Cleaned Quotes"

ws.append(["ID", "Quote Text", "Author", "Tags"])

url = "https://quotes.toscrape.com/"

try:
    print("Getting Data and Cleaning....")
    response = requests.get(url, timeout=10)

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")
        quotes = soup.find_all("div", class_="quote")

        for index, item in enumerate(quotes, 1):

            text_el = item.find("apan", class_="text")
            author_el = item.find("small", class_="author")

            cleaned_text = text_el.text.strip().strip('"').strip('"').strip('"') if text_el else "N/A"
            cleaned_author = author_el.text.strip() if author_el else "N/A"


            tags_elements = item.find_all("a", class_="tag")
            tags_list = [t.text.strip() for t in tags_elements if t and t.text]
            cleaned_tags = ", ".join(tags_list)

            ws.append([index, cleaned_text, cleaned_author, cleaned_tags])

        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_file = os.path.join(script_dir, "cleaned_quotes.xlsx")

        wb.save(output_file)
        print("\n-------------------------------------------------")
        print(f" Successful. Clean Data save to: '{output_file}'")

    else:
        print(f" Failed to connecct. Status Code: {response. status_code}")
except Exception as e:
    print(f" Falseting: {e}")                