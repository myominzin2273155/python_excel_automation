import openpyxl

wb = openpyxl.Workbook()

sheet = wb.active
sheet.title = "SalesData"

sheet["A1"] = "Item"
sheet["B1"] = "Quantity"
sheet["C1"] = "Price"

sheet.append(["Apple", 10, 1500])
sheet.append(["Orange", 5, 2000])
sheet.append(["Banana", 12, 500])

wb.save("first_automation.xlsx")

print("Excel File အောင်မြင်စွာ ဖန်တီးပြီးပါပြီခင်ဗျာ!") 