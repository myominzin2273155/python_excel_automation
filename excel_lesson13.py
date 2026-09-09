import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os

# 1. အီးမေးလ်၏ အခြေခံ အချက်အလက်များ သတ်မှတ်ခြင်း
sender_email = "your_email@gmail.com"
receiver_email = "client_email@gmail.com"
subject = "Daily Sales Report Automation"
body = "မင်္ဂလာပါခင်ဗျာ၊ ယနေ့အတွက် အလိုအလျောက် ထုတ်ယူထားသော Sales Report PDF ဖိုင်ကို ပူးတွဲ ပေးပို့အပ်ပါသည်။"

# 2. Email Message Object ဖန်တီးခြင်း
msg = MIMEMultipart()
msg["From"] = sender_email
msg["To"] = receiver_email
msg["Subject"] = subject
msg.attach(MIMEText(body, "plain"))

# 3. လက်ရှိ Folder ထဲရှိ PDF ဖိုင်ကို အတိအကျ ရှာယူခြင်း
script_dir = os.path.dirname(os.path.abspath(__file__))
pdf_path = os.path.join(script_dir, "freelance_final_report.pdf")

if os.path.exists(pdf_path):
    with open(pdf_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(pdf_path))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(pdf_path)}"'
        msg.attach(part)
    print("PDF File ကို အီးမေးလ်ထဲသို့ အောင်မြင်စွာ ပူးတွဲပြီးပါပြီ။")
else:
    # Folder ထဲမှာ တခြား pdf ဖိုင်ရှိမရှိ ရှာဖွေပေးခြင်း
    pdf_files = [f for f in os.listdir(script_dir) if f.endswith('.pdf')]
    if pdf_files:
        found_pdf = os.path.join(script_dir, pdf_files[0])
        with open(found_pdf, "rb") as f:
            part = MIMEApplication(f.read(), Name=pdf_files[0])
            part["Content-Disposition"] = f'attachment; filename="{pdf_files[0]}"'
            msg.attach(part)
        print(f"တွေ့ရှိထားသော '{pdf_files[0]}' ကို အီးမေးလ်ထဲသို့ ပူးတွဲလိုက်ပါပြီ။")
    else:
        print("သတိပေးချက်: Folder ထဲတွင် PDF ဖိုင် လုံးဝ ရှာမတွေ့ပါ။")

# 4. အီးမေးလ် ပေးပို့သည့် တည်ဆောက်ပုံ (Logic Preview)
print("Email Sending Logic Ready! (Real Mail ပို့ရန် SMTP Server ချိန်ဆရန် လိုအပ်ပါသည်)")