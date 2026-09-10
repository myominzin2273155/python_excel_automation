import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import os
import time

wb = Workbook()
ws = wb.active
ws.title = "All Pages Quotes"

ws.append(["No", "Quote Text", "Author", "Page Number"])

base_url = "https://quotes.toscrape.com"
next_page_url = "/page/1/"
page_number = 1
total_quotes_count = 0

try:
    while next_page_url:
        full_url = base_url + next_page_url
        print(f" page ({page_number}) from Data pull out: {full_url}")

        response = requests.get(full_url, timeout=10)

        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "html.parser")
            quotes = soup.find_all("div", class_="quote")

            for item in quotes:
                total_quotes_count += 1
                text = item.find("span", class_="text").text
                author = item.find("small", class_="author").text
                ws.append([total_quotes_count, text, author, f"page {page_number}"])

            next_btn = soup.find("li", class_="next")
            if next_btn and next_btn.find("a"):
                next_page_url = next_btn.find("a")["href"]
                page_number += 1
                time.sleep(1)
            else:
                next_page_url = None
        else:
            print(f" page {page_number} no success join.") 
            break

    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_file = os.path.join(script_dir, "all_pages_quotes.xlsx")

    wb.save(output_file)
    print("\n--------------------------------------------")
    print(f" Successful. Sum page ({page_number}) from sar so ({total_quotes_count}) pull out")
    print(f" Save as files: '{output_file}")

except Exception as e:
    print(f" Falseting: {e}")

