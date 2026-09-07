import openpyxl

wb = openpyxl.load_workbook("first_automation.xlsx")
sheet = wb["SalesData"]

sheet["D1"] = "Total"

for row in range(2, sheet.max_row + 1):
    quantity = sheet[f"B{row}"].value
    price = sheet[f"C{row}"].value

    total = quantity * price
    sheet[f"D{row}"] = total
wb.save("calculated_sales.xlsx")
print("Data များကို အောင်မြင်စွာ တွက်ချက်ပြီး Calculated_sales.xlsx ထဲသို့ သိမ်းဆည်းလိုက်ပါပြီခင်ဗျာ!")    