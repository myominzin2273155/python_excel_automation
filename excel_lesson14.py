import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os

sender_email = "your_email@gmail.com"
app_password = "xxxx xxxx xxxx xxxx"
receiver_email = "your_email@gmail.com"

msg = MIMEMultipart()
msg["From"] = sender_email
msg["To"] = receiver_email
msg["Subject"] = "Automated Daily Report - Freelance Pipeline"

body = " မင်္ဂလာပါခင်ဗျာ Python Automation မှ အလိုအလျောက် ထုတ်ယူ ပေးပို့သော Report ဖြစ်ပါသည်။"
msg.attach(MIMEText(body, "plain"))

script_dir = os.path.dirname(os.path.abspath(__file__))
pdf_path = os.path.join(script_dir, "final_sales_report.pdf")

if os.path.exists(pdf_path):
    with open(pdf_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(pdf_path))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(pdf_path)}"'
        msg.attach(part)
    print("PDF File ကို အီးမေးလ်ထဲသို့ ပူးတွဲလိုက်ပါပြီ။")
try:
    print("Gmail SMTP Server သို့ ချိတ်ဆက်နေပါသည်...")
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender_email, app_password)

    text = msg.as_string()
    server.sendmail(sender_email, receiver_email, text)
    server.quit()

    print(" အီးမေးလ်ကို တကယ့် မေးလ်ဘောက်စ်ထဲသို့ အောင်မြင်စွာ ပေးပို့လိုက်ပါပြီ")

except Exception as e:
    print(f" အီးမေးလ် ပို့ရာတွင် အမှားဖြစ်ပေါ်ပါသည်: {e}")    
        