import os
import openpyxl
from openpyxl.chart import BarChart, Reference
import win32com.client

wb = openpyxl.load_workbook("master_all_folder_sales.xlsx")
sheet = wb.active

new_wb = openpyxl.Workbook()
new_sheet = new_wb.active
new_sheet.title = "Summary"

header = [cell.value for cell in sheet[1]]
new_sheet.append(header)

seen_rows = set()
for row in sheet.iter_rows(min_row=2, values_only=True):
    if not any(row):
        continue
    if row not in seen_rows:
        seen_rows.add(row)
        new_sheet.append(row)

chart = BarChart()
chart.type = "col"
chart.style = 10
chart.title = "Final Sales Performance"
chart.y_axis.title = "Amount"
chart.x_axis.title = "Items"

data = Reference(new_sheet, min_col=2, min_row=1, max_col=2, max_row=new_sheet.max_row)
cats = Reference(new_sheet, min_col=1, min_row=2, max_row=new_sheet.max_row)

chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
new_sheet.add_chart(chart, "E2")

cleaned_excel = os.path.abspath("freelance_final_report.xlsx")
new_wb.save(cleaned_excel)

pdf_output = os.path.abspath("freelance_final_report.xlsx")
excel_app = win32com.client.Dispatch("Excel.Application")
excel_app.Visible = False

try:
    doc = excel_app.Workbooks.Open(cleaned_excel)
    doc.ExportAsFixedFormat(0, pdf_output)
    doc.Close()
    print("Freelance Project Pipeline စက္ကန့်ပိုင်းအတွင်း အောင်မြင်ပြီး")

except Exception as e:
    print(f"Error ဖြစ်ပေါ်ပါသည် : {e}")

finally:
    excel_app.Quit()    