import win32com.client
import os

excel_file = os.path.abspath("sales_report_with_chart.xlsx")
pdf_file = os.path.abspath("final_sales_report.pdf")

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

try:
    wb = excel.Workbooks.Open(excel_file)
    wb.ExportAsFixedFormat(0, pdf_file)
    wb.Close()
    print("Excel File ကို pdf အဖြစ် အောင်မြင်စွာ ပြောင်းလဲပြီးပါပြီ")

except Exception as e:
    print(f"Error ဖြစ်ပေါ်ပါသည်: {e}")

finally:
    excel.Quit()

