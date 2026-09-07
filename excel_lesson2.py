import openpyxl

wb = openpyxl.load_workbook("first_automation.xlsx")

sheet = wb["SalesData"]

print("--- Cell တစ်ခုချင်းစီ ဖတ်ခြင်း ---")
print("Cell A1:", sheet["A1"].value)
print("Cell B2(Apple Quantity):", sheet["B2"].value)
print("Cell C2 (Apple Price):", sheet["C2"].value)

print("\n--- Row အားလုံးကို Loop ပတ်၍ ဖတ်ခြင်း ---")

for row in sheet.iter_rows(values_only=True):
    print(row)