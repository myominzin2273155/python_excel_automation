import os
from bs4 import BeautifulSoup
from deep_translator import GoogleTranslator
from openpyxl import Workbook
import requests

translator = GoogleTranslator(source="auto", target="my")

wb = Workbook()
ws = wb.active
ws.title = "Cleaned Quotes"

ws.append(["ID", "Quote Text (Eng)", "Quote Text (MM)", "Author", "Tags"])

url = "https://quotes.toscrape.com/"

try:
    print("Getting Data, Cleaning, and Translating to Myanmar...")
    response = requests.get(url, timeout=10)

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")
        quotes = soup.find_all("div", class_="quote")

        for index, item in enumerate(quotes, 1):
            # ပြင်ဆင်ချက် ၁: "apan" ကို "span" သို့ ပြင်ဆင်ထားပါသည်
            text_el = item.find("span", class_="text")
            author_el = item.find("small", class_="author")

            cleaned_text = (
                text_el.text.strip().strip('"').strip("“").strip("”")
                if text_el
                else "N/A"
            )
            cleaned_author = author_el.text.strip() if author_el else "N/A"

            # စာသားရှိပါက မြန်မာလို ဘာသာပြန်မည်
            translated_mm = (
                translator.translate(cleaned_text)
                if cleaned_text != "N/A"
                else "N/A"
            )

            tags_elements = item.find_all("a", class_="tag")
            tags_list = [t.text.strip() for t in tags_elements if t and t.text]
            cleaned_tags = ", ".join(tags_list)

            # Excel ထဲသို့ Row တစ်ခုချင်းစီ ထည့်မည်
            ws.append(
                [
                    index,
                    cleaned_text,
                    translated_mm,
                    cleaned_author,
                    cleaned_tags,
                ]
            )

        # ပြင်ဆင်ချက် ၂: File save အပိုင်းကို Indentation စနစ်တကျ ပြန်ညှိထားပါသည်
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_file = os.path.join(script_dir, "cleaned_quotes.xlsx")

        wb.save(output_file)
        print("\n-------------------------------------------------")
        print(
            f" Successful. Clean & Myanmar Translated Data saved to:"
            f" '{output_file}'"
        )

    else:
        print(f" Failed to connect. Status Code: {response.status_code}")

except Exception as e:
    print(f" Error: {e}")