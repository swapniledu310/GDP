import os
import smtplib
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo


EXCEL_FILE = Path("IPO_Index_GMP_Dashboard.xlsx")

gmail_username = os.environ["GMAIL_USERNAME"]
gmail_app_password = os.environ["GMAIL_APP_PASSWORD"]
email_to = os.environ["EMAIL_TO"]


if not EXCEL_FILE.exists():
    raise FileNotFoundError(
        f"Excel report was not generated: {EXCEL_FILE.resolve()}"
    )


today_ist = datetime.now(
    ZoneInfo("Asia/Kolkata")
).strftime("%d-%b-%Y")


message = EmailMessage()

message["From"] = gmail_username
message["To"] = email_to
message["Subject"] = f"Daily IPO GMP Dashboard – {today_ist}"

message.set_content(
    f"""Hello,

Please find attached the Daily IPO GMP Dashboard for {today_ist}.

The report contains:
- Latest IPO GMP information
- IPO status and segment analysis
- Six pivot-analysis tables
- Interactive Excel dashboard
- GMP and status filters
- Dashboard charts

Source: https://ipoindex.in/ipo-gmp/

Regards,
Swapnil
"""
)


with EXCEL_FILE.open("rb") as file:
    message.add_attachment(
        file.read(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=EXCEL_FILE.name,
    )


with smtplib.SMTP_SSL(
    "smtp.gmail.com",
    465,
    timeout=60,
) as smtp:
    smtp.login(
        gmail_username,
        gmail_app_password,
    )

    smtp.send_message(message)


print(
    f"Email successfully sent to {email_to} "
    f"with attachment {EXCEL_FILE.name}"
)
