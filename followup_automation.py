import smtplib
import imaplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
import time
import traceback
from googleapiclient.discovery import build
import json
import os

# === SMTP/IMAP Credentials ===
SMTP_SERVER = "mail.b2bgrowthexpo.com"
SMTP_PORT = 587
SMTP_EMAIL = "nagendra@b2bgrowthexpo.com"
SMTP_PASSWORD = "tdA^87%+.3$3"

IMAP_SERVER = "mail.b2bgrowthexpo.com"
IMAP_PORT = 143
IMAP_EMAIL = SMTP_EMAIL
IMAP_PASSWORD = SMTP_PASSWORD

SENDER_NAME = "Nagendra Mishra"

# === HTML Email Template ===
EMAIL_TEMPLATE = """
<html>
  <body style="font-family: Arial, sans-serif; font-size: 15px; color: #333; background-color: #ffffff; padding: 20px;">
    <div style="text-align: center; margin-bottom: 20px;">
      <img src="https://iili.io/FogC9l2.jpg" alt="B2B Growth Expo" style="max-width: 400px; height: auto;" />
    </div>
    <p>Hi {%name%},</p>
    <p>{%body%}</p>
    <p>
      If you would like to schedule a meeting with me at your convenient time,<br>
      please use the link below:<br>
      <a href="https://tidycal.com/nagendra/b2b-discovery-call" target="_blank">https://tidycal.com/nagendra/b2b-discovery-call</a>
    </p>
    <p style="margin-top: 30px;">
      Thanks & Regards,<br>
      <strong>Nagendra Mishra</strong><br>
      Director | B2B Growth Hub<br>
      Mo: +44 7913 027482<br>
      Email: <a href="mailto:nagendra@b2bgrowthhub.com">nagendra@b2bgrowthhub.com</a><br>
      <a href="https://www.b2bgrowthhub.com" target="_blank">www.b2bgrowthhub.com</a>
    </p>
    <p style="font-size: 13px; color: #888;">
      If you don’t want to hear from me again, please let me know.
    </p>
  </body>
</html>
"""

# === Authenticate Google Sheets ===
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file('/etc/secrets/google-credentials.json', scopes=SCOPES)
sheets_api = build("sheets", "v4", credentials=creds)
gc = gspread.authorize(creds)
sheet = gc.open("Sales-sheet-automation-test").worksheet("Sales")

# === Follow-up Templates ===
FOLLOWUP_EMAILS = [
    "This is Nagendra from B2B Growth Expo. Thank you for expressing interest in exhibiting at our upcoming {%show%}. I'd love to schedule a quick call to understand your requirements better. Could you let me know a suitable time for a short conversation?",
    "Since I haven’t heard back, I’m sharing our Exhibitor Pitch Deck to help you make a more informed decision.<br>Here is the link: <a href=\"{%pitch_deck_url%}\" target=\"_blank\">Pitch Deck</a><br>Feel free to reach out if you have any questions.",
    "Just checking in—were you able to go through the Exhibitor Pitch Deck I shared earlier?",
    "I understand things can get busy. I'd appreciate it if you could take a moment to let me know your thoughts when you have a chance."
]

FOLLOWUP_SUBJECTS = [
    "Let's Discuss Your Participation at the Upcoming {%show%}",
    "Exhibitor Pitch Deck Inside – Still Interested?",
    "Did You Get a Chance to Review the Pitch Deck?",
    "Just Checking In – Any Thoughts?"
]

FINAL_EMAIL = (
    "I completely understand if your initial interest was out of curiosity. "
    "If you're no longer interested in exhibiting, that’s absolutely fine—just let me know with a simple 'Yes' or 'No' "
    "so we don’t keep reaching out unnecessarily."
)

def send_email(to_email, subject, body, name=""):
    print(f"Preparing to send email to: {to_email}")
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"{SENDER_NAME} <{SMTP_EMAIL}>"
    msg["To"] = to_email
    html_body = EMAIL_TEMPLATE.replace("{%name%}", name).replace("{%body%}", body)
    msg.attach(MIMEText(html_body, "html"))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_EMAIL, SMTP_PASSWORD)
            server.sendmail(SMTP_EMAIL, to_email, msg.as_string())
        print(f"✅ Email sent to {to_email}")
    except Exception as e:
        print(f"❌ SMTP Error while sending to {to_email}: {e}")

    try:
        imap = imaplib.IMAP4_SSL(IMAP_SERVER)
        imap.login(SMTP_EMAIL, SMTP_PASSWORD)
        imap.append("INBOX.Sent", '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        imap.logout()
    except Exception as e:
        print(f"❌ IMAP Error while saving to Sent folder for {to_email}: {e}")

def get_reply_emails():
    print("Checking for new replies in inbox...")
    replied = set()
    try:
        with imaplib.IMAP4_SSL(IMAP_SERVER) as mail:
            mail.login(IMAP_EMAIL, IMAP_PASSWORD)
            mail.select("inbox")
            status, messages = mail.search(None, 'UNSEEN')
            for num in messages[0].split():
                _, data = mail.fetch(num, "(RFC822)")
                msg = email.message_from_bytes(data[0][1])
                from_addr = email.utils.parseaddr(msg["From"])[1].lower().strip()
                replied.add(from_addr)
    except Exception as e:
        print(f"❌ IMAP Error while checking replies: {e}")
    print(f"✅ Found {len(replied)} new replies.")
    return replied

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return {
        "red": int(hex_color[0:2], 16) / 255,
        "green": int(hex_color[2:4], 16) / 255,
        "blue": int(hex_color[4:6], 16) / 255
    }

def set_row_color(sheet, row_number, color_hex):
    print(f"Coloring row {row_number} with color {color_hex}")
    try:
        sheet_format = {
            "requests": [{
                "repeatCell": {
                    "range": {
                        "sheetId": sheet._properties['sheetId'],
                        "startRowIndex": row_number - 1,
                        "endRowIndex": row_number,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColorStyle": {
                                "rgbColor": hex_to_rgb(color_hex)
                            }
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColorStyle"
                }
            }]
        }
        sheet.spreadsheet.batch_update(sheet_format)
    except Exception as e:
        print(f"❌ Google Sheets Error while coloring row {row_number}: {e}")

def get_row_background_color(sheet_id, sheet_name, row_number):
    try:
        range_ = f"{sheet_name}!A{row_number}"
        result = sheets_api.spreadsheets().get(
            spreadsheetId=sheet_id,
            ranges=[range_],
            fields="sheets.data.rowData.values.effectiveFormat.backgroundColor"
        ).execute()

        cell_format = result['sheets'][0]['data'][0]['rowData'][0]['values'][0]['effectiveFormat']['backgroundColor']
        rgb = (
            int(cell_format.get('red', 0) * 255),
            int(cell_format.get('green', 0) * 255),
            int(cell_format.get('blue', 0) * 255)
        )
        print(f"Row {row_number} color fetched: RGB{rgb}")
        return rgb
    except Exception as e:
        print(f"❌ Error getting background color for row {row_number}: {e}")
        return None

def process_replies():
    print("Processing replies...")
    try:
        data = sheet.get_all_records()
        replied_emails = get_reply_emails()
        for idx, row in enumerate(data, start=2):
            if not any(row.values()):
                continue

            email_addr = row.get("Email", "").lower().strip()
            rgb = get_row_background_color(sheet.spreadsheet.id, sheet.title, idx)

            if rgb:
                r, g, b = rgb

                if r > 240 and g > 240 and b > 240:
                    print(f"Row {idx} is white (no color), processing.")
                if r > 180 and g < 100 and b < 100:
                    print(f"Row {idx} is red, skipping.")
                    continue
                if r < 100 and g > 180 and b < 100:
                    print(f"Row {idx} is green, skipping.")
                    continue
                if abs(r - 255) < 10 and abs(g - 255) < 10 and b < 50:
                    print(f"Row {idx} is yellow, skipping.")
                    continue

            if not email_addr:
                continue

            if email_addr in replied_emails and row.get("Reply Status", "") != "Replied":
                print(f"Marking row {idx} ({email_addr}) as Replied.")
                sheet.update_cell(idx, 7, "Replied")
                set_row_color(sheet, idx, "#FFFF00")  # Yellow

    except Exception as e:
        print("❌ Error in processing replies:", e)

def process_followups():
    print("Processing follow-up emails...")
    try:
        data = sheet.get_all_records()
        today = datetime.today().strftime('%Y-%m-%d')

        for idx, row in enumerate(data, start=2):
            try:
                print(f"\nRow {idx}: Checking {row.get('Email')}")
                if not any(row.values()):
                    continue
                rgb = get_row_background_color(sheet.spreadsheet.id, sheet.title, idx)
                if rgb:
                    r, g, b = rgb
                    if r > 240 and g > 240 and b > 240:
                        print(f"Row {idx} is white (no color), processing.")
                    elif r > 180 and g < 100 and b < 100:
                        print(f"Row {idx} is red, skipping.")
                        continue
                    elif r < 100 and g > 180 and b < 100:
                        print(f"Row {idx} is green, skipping.")
                        continue
                    elif abs(r - 255) < 10 and abs(g - 255) < 10 and b < 50:
                        print(f"Row {idx} is yellow, skipping.")
                        continue

                email_addr = row.get("Email", "").lower().strip()
                if not email_addr:
                    continue

                name = row.get("First_Name", "").strip()
                count = int(row.get("Follow-Up Count") or 0)
                last_date = row.get("Last Follow-Up Date", "")
                reply_status = row.get("Reply Status", "").strip()

                if reply_status in ["Replied", "No Reply After 4"]:
                    continue

                if last_date:
                    last_dt = datetime.strptime(last_date, "%Y-%m-%d")
                    if (datetime.now() - last_dt).total_seconds() < 86400:
                        continue

                if count >= 4:
                    send_email(email_addr, "Should I Close Your File?", FINAL_EMAIL, name=name)
                    sheet.update_cell(idx, 7, "No Reply After 4 Followups")
                    set_row_color(sheet, idx, "#FF0000")
                    continue

                followup_text = FOLLOWUP_EMAILS[count].replace("{%name%}", name)
                subject = FOLLOWUP_SUBJECTS[count]

                if count == 0:
                    show = row.get("Show", "").strip()
                    if not show:
                        continue
                    followup_text = followup_text.replace("{%show%}", show)
                    subject = subject.replace("{%show%}", show)
                elif count == 1:
                    url = row.get("Pitch Deck URL", "").strip()
                    if not url:
                        continue
                    followup_text = followup_text.replace("{%pitch_deck_url%}", url)

                send_email(email_addr, subject, followup_text, name=name)
                sheet.update_cell(idx, 5, str(count + 1))
                sheet.update_cell(idx, 6, today)
                sheet.update_cell(idx, 7, "Pending")

                if (idx - 1) % 3 == 0:
                    print("Sleeping for 3 seconds...")
                    time.sleep(3)

            except Exception as e:
                print(f"❌ Error on row {idx}: {e}")
    except Exception as e:
        print("❌ Error in processing followups:", e)

# === Entry Point ===
if __name__ == "__main__":
    print("🚀 Sales follow-up automation started...")
    next_followup_check = time.time()
    
    while True:
        try:
            print("\n--- Checking for replies ---")
            process_replies()  # Every 15 minutes

            current_time = time.time()

            if current_time >= next_followup_check:
                print("\n--- Sending follow-up emails ---")
                process_followups()  # Every 60 minutes
                next_followup_check = current_time + 3600  # Set next follow-up in 60 min

        except Exception:
            print("❌ Fatal error:")
            traceback.print_exc()

        time.sleep(900)  # Sleep 15 minutes before next reply check
