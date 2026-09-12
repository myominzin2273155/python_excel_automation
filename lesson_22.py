import os
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------------------------------------------------------
# ၁။ Auto-fit Column Width (စာသားအရှည်အတိုင်း ကွက်နံပါတ် ညှိပေးခြင်း)
# ---------------------------------------------------------
def auto_fit_columns(ws):
    """Excel ဇယားကွက်များ မကျဉ်းစေရန် စာသားအရှည်အတိုင်း Width ကို အလိုအလျောက် ညှိပေးသော Function"""
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter  # Column နာမည် (A, B, C...) ယူခြင်း
        
        for cell in col:
            if cell.value:
                # စာသားအရှည်ကို တိုင်းတာခြင်း
                cell_len = len(str(cell.value))
                if cell_len > max_len:
                    max_len = cell_len
        
        # စာသား မရှုပ်ရအောင် ဘေးဘက်တွင် ကွက်လပ် ၃ ကွက် စာ ပိုပေးခြင်း (Maximum Width ကို ၆၀ သတ်မှတ်ထားသည်)
        ws.column_dimensions[col_letter].width = min(max_len + 3, 60)

# ---------------------------------------------------------
# ၂။ Zebra Striping & Formatting (တစ်ကြောင်းခြား အရောင်ဖြည့်ခြင်း)
# ---------------------------------------------------------
def apply_zebra_styling(ws):
    """Data များကို ဖတ်ရလွယ်ကူစေရန် တစ်ကြောင်းခြား အရောင်နု ဖြည့်ပေးသော Function"""
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    
    # Alternate Row Color (အပြာနုရောင် အပျော့လေး)
    zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    
    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3")
    )

    # Header အလှဆင်ခြင်း
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Data Rows အလှဆင်ခြင်း (မမ ဂဏန်း Row များကို အရောင် ဖြည့်ပေးမည်)
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
        for cell in row:
            cell.border = thin_border
            
            # စာကြောင်း ရဲဲ့ Row ID က စုံဂဏန်း ဖြစ်ပါက အပြာနုရောင် ထည့်မည်
            if row_idx % 2 == 0:
                cell.fill = zebra_fill

            # Alignments
            if cell.column == 1:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

# ---------------------------------------------------------
# ၃။ Main Execution (အဓိက အလုပ်လုပ်သည့်နေရာ)
# ---------------------------------------------------------
def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Quotes Data"

    # Header ရေးသားခြင်း
    ws.append(["ID", "Quote Text (Eng)", "Author", "Tags"])

    # Scraping အပိုင်း (စာမျက်နှာ ၁ တစ်ခုတည်းကိုသာ စမ်းသပ်မည်)
    url = "https://quotes.toscrape.com"
    response = requests.get(url)
    soup = BeautifulSoup(response.text, "html.parser")
    quotes = soup.find_all("div", class_="quote")

    for i, item in enumerate(quotes, start=1):
        text = item.find("span", class_="text").text.strip().strip('"')
        author = item.find("small", class_="author").text.strip()
        tags = ", ".join([t.text.strip() for t in item.find_all("a", class_="tag")])
        
        ws.append([i, text, author, tags])

    # Styling Function များကို လှမ်းခေါ်သုံးခြင်း
    auto_fit_columns(ws)
    apply_zebra_styling(ws)

    # Excel ဖိုင် သိမ်းဆည်းခြင်း
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_file = os.path.join(script_dir, "lesson_22_styled_quotes.xlsx")
    wb.save(output_file)
    print("Lesson 22 Styled Data Saved Successfully!")

if __name__ == "__main__":
    main()