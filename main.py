"""
Gmail Mail Merge
-----------------
Sends the same email draft to everyone in an Excel file, personalizing with
their Name if you provide one. Safe to re-run: it skips anyone already
marked as "Sent" so you only reach new rows each time.

SETUP (one-time):
1. Turn on 2-Step Verification on your Google account:
   https://myaccount.google.com/signinoptions/two-step-verification
2. Create an App Password:
   https://myaccount.google.com/apppasswords
   (choose "Mail" as the app) -> copy the 16-character password it gives you
3. Fill in the CONFIG section below.
4. Your Excel file needs a column called "Email" (required) and optionally
   "Name" (used for personalization). A "Sent" column will be added/updated
   automatically -- don't remove it once it appears.

RUN:
    pip install openpyxl --break-system-packages
    python gmail_mail_merge.py
"""

import os
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import openpyxl

# ----------------------- CONFIG: edit these ----------------------------
GMAIL_ADDRESS = "you@gmail.com"        # your Gmail address
GMAIL_APP_PASSWORD = "xxxx xxxx xxxx xxxx"  # the 16-char App Password (no regular password!)

EXCEL_PATH = "recipients.xlsx"         # path to your Excel file
EMAIL_COLUMN = "Email"                 # column header with email addresses
NAME_COLUMN = "Name"                   # column header with names (optional; set to None if you don't have one)
SENT_COLUMN = "Sent"                   # column that tracks who's already been emailed

SUBJECT = "Your subject line here"

# Use {name} anywhere you want it swapped in. Leave as plain text otherwise.
BODY_TEMPLATE = """Hi {name},

This is the email body. Replace this with your actual message.

Best,
Your Name
"""

# Path to a PDF to attach to every email. Set to None to send with no attachment.
PDF_ATTACHMENT = "attachment.pdf"
# -------------------------------------------------------------------------


def send_email(to_address, subject, body, attachment_path=None):
    msg = MIMEMultipart()
    msg["From"] = GMAIL_ADDRESS
    msg["To"] = to_address
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    if attachment_path:
        with open(attachment_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        filename = os.path.basename(attachment_path)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        msg.attach(part)

    context = ssl.create_default_context()
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls(context=context)
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, to_address, msg.as_string())


def main():
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]

    if EMAIL_COLUMN not in headers:
        raise ValueError(f"Couldn't find a '{EMAIL_COLUMN}' column in {EXCEL_PATH}")

    email_col = headers.index(EMAIL_COLUMN) + 1
    name_col = headers.index(NAME_COLUMN) + 1 if NAME_COLUMN and NAME_COLUMN in headers else None

    if SENT_COLUMN in headers:
        sent_col = headers.index(SENT_COLUMN) + 1
    else:
        sent_col = len(headers) + 1
        ws.cell(row=1, column=sent_col, value=SENT_COLUMN)

    sent_count = 0
    skipped_count = 0

    for row in range(2, ws.max_row + 1):
        email = ws.cell(row=row, column=email_col).value
        already_sent = ws.cell(row=row, column=sent_col).value

        if not email:
            continue
        if already_sent:
            skipped_count += 1
            continue

        name = ws.cell(row=row, column=name_col).value if name_col else ""
        first_name = str(name).split()[0] if name else "there"

        body = BODY_TEMPLATE.format(name=first_name)

        try:
            send_email(email, SUBJECT, body, PDF_ATTACHMENT)
            ws.cell(row=row, column=sent_col, value=f"Sent {datetime.now():%Y-%m-%d %H:%M}")
            sent_count += 1
            print(f"Sent to {email}")
        except Exception as e:
            ws.cell(row=row, column=sent_col, value=f"FAILED: {e}")
            print(f"Failed to send to {email}: {e}")

    wb.save(EXCEL_PATH)
    print(f"\nDone. Sent: {sent_count}, Skipped (already sent): {skipped_count}")


if __name__ == "__main__":
    main()