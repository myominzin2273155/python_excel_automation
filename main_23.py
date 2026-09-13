import os
from openpyxl import Workbook
from scraper import fetch_quotes_data
from excel_formatter import auto_fit_columns, apply_zebra_styling

def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Modular Quotes Data"

    ws.append(["ID", "Quote Text (Eng)", "Author", "Tags"])

    quotes_data = fetch_quotes_data()

    for row in quotes_data:
        ws.append(row)

    auto_fit_columns(ws)
    apply_zebra_styling(ws)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_file = os.path.join(script_dir, "lesson_23_modular_output.xlsx")
    wb.save(output_file)
    print("Lesson 23 Modular Automation Pipeline Completed Successfully!")

if __name__ == "__main__":
    main()        