import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import os
from deep_translator import GoogleTranslator

# Google Translator ပြင်ဆင်ခြင်း (Auto Detect မှ မြန်မာစာသို့)
translator = GoogleTranslator(source='auto', target='my')

wb = Workbook()
ws = wb.active
ws.title = "Cleaned Quotes"

# Header ထဲတွင် မြန်မာစာ Column ပါဝင်သည်
ws.append(["ID", "Quote Text (Eng)", "Quote Text (MM)", "Author", "Tags"])

url = "https://quotes.toscrape.com/"

try:
    print("Getting Data, Cleaning, and Translating to Myanmar...")
    response = requests.get(url, timeout=10)
    
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")
        quotes = soup.find_all("div", class_="quote")
        
        for index, item in enumerate(quotes, 1):
            text_el = item.find("span", class_="text")
            author_el = item.find("small", class_="author")
            
            cleaned_text = text_el.text.strip().strip('"').strip('“').strip('”') if text_el else "N/A"
            cleaned_author = author_el.text.strip() if author_el else "N/A"
            
            # မြန်မာလို အလိုအလျောက် ဘာသာပြန်ခြင်း
            translated_mm = translator.translate(cleaned_text) if cleaned_text != "N/A" else "N/A"
            print(f"[{index}/10] Translating: {cleaned_author}")
            
            tags_elements = item.find_all("a", class_="tag")
            tags_list = [t.text.strip() for t in tags_elements if t and t.text]
            cleaned_tags = ", ".join(tags_list)
            
            # Excel ထဲ ပူးတွဲ သိမ်းဆည်းခြင်း
            ws.append([index, cleaned_text, translated_mm, cleaned_author, cleaned_tags])
            
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_file = os.path.join(script_dir, "cleaned_quotes.xlsx")
        
        wb.save(output_file)
        print("\n--------------------------------------------------")
        print(f"Success! Cleaned & Myanmar Translated Data saved to: '{output_file}'")
        
    else:
        print(f"Failed to connect. Status Code: {response.status_code}")

except Exception as e:
    print(f"Error: {e}")