import csv
import openpyxl

csv_data = [
    ["Item", "Quantity", "Price"],
    ["Apple", "10", "1500"],
    ["Orange", "INVALID", "2000"],
    ["Banana", "12", "500"]
]

with open("raw_sales.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerows(csv_data)

wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "ValidSales"
sheet.append(["Item", "Quantity", "Price", "Total"])

with open("raw_sales.csv", "r", encoding="utf-8") as f:
    reader = csv.reader(f)
    next(reader) 

    for row in reader:
        try:
            item = row[0]
            qty = int(row[1]) 
            price = int(row[2])
            total = qty * price
            sheet.append([item, qty, price, total])
        except ValueError:
            print(f"သတိပေးချက်: '{row[0]}' ၏ Data မှားယွင်းနေသဖြင့် ကျော်ခွလိုက်ပါသည်- {row}")
            continue

wb.save("cleaned_csv_sales.xlsx")
print("CSV Data များကို Error စစ်ဆေးပြီး Excel သို့ အောင်မြင်စွာ ပြောင်းလဲပြီးပါပြီ။")