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
SMTP_PASSWORD = "D@shwood0404"

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
      Email: <a href="mailto:nagendra@b2bgrowthexpo.com">nagendra@b2bgrowthexpo.com</a><br>
      <a href="https://www.b2bgrowthexpo.com" target="_blank">www.b2bgrowthexpo.com</a>
    </p>
    <p style="font-size: 13px; color: #888;">
      If you don’t want to hear from me again, please let me know.
    </p>
  </body>
</html>
"""

# === Authenticate Google Sheets ===
print("🔐 Authenticating Google Sheets...", flush=True)
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file('/etc/secrets/google-credentials.json', scopes=SCOPES)
sheets_api = build("sheets", "v4", credentials=creds)
gc = gspread.authorize(creds)
sheet = gc.open("Expo-Sales-Management").worksheet("exhibitors-1")
print("✅ Google Sheets authenticated and worksheet loaded.", flush=True)

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

# === Email Sending ===
def send_email(to_email, subject, body, name=""):
    print(f"✉️ Preparing to send email to: {to_email}", flush=True)
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"{SENDER_NAME} <{SMTP_EMAIL}>"
    msg["To"] = to_email
    html_body = EMAIL_TEMPLATE.replace("{%name%}", name).replace("{%body%}", body)
    msg.attach(MIMEText(html_body, "html"))

    try:
        print(f"🔄 Connecting to SMTP server {SMTP_SERVER}:{SMTP_PORT}...", flush=True)
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            print("🔑 Logging into SMTP server...", flush=True)
            server.login(SMTP_EMAIL, SMTP_PASSWORD)
            server.sendmail(SMTP_EMAIL, to_email, msg.as_string())
        print(f"✅ Email sent successfully to {to_email}", flush=True)
    except Exception as e:
        print(f"❌ SMTP Error while sending to {to_email}: {e}", flush=True)

    try:
        print(f"🔄 Saving email to Sent folder via IMAP...", flush=True)
        imap = imaplib.IMAP4_SSL(IMAP_SERVER)
        imap.login(SMTP_EMAIL, SMTP_PASSWORD)
        imap.append("INBOX.Sent", '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        imap.logout()
        print(f"✅ Email saved in Sent folder for {to_email}", flush=True)
    except Exception as e:
        print(f"❌ IMAP Error while saving to Sent folder for {to_email}: {e}", flush=True)

# === Fetch replies from inbox ===
def get_reply_emails():
    print("🔍 Checking for new replies in inbox...", flush=True)
    replied = set()
    try:
        try:
            print("🔐 Trying IMAP SSL connection on port 993...", flush=True)
            mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
            mail.login(IMAP_EMAIL, IMAP_PASSWORD)
            print("✅ Connected via IMAP SSL (993)", flush=True)
        except Exception as ssl_err:
            print(f"⚠️ SSL connection failed: {ssl_err}", flush=True)
            print("🔄 Trying STARTTLS on port 143...", flush=True)
            mail = imaplib.IMAP4(IMAP_SERVER, 143)
            mail.starttls()
            mail.login(IMAP_EMAIL, IMAP_PASSWORD)
            print("✅ Connected via IMAP STARTTLS (143)", flush=True)

        mail.select("INBOX")
        status, messages = mail.search(None, 'UNSEEN')
        print(f"📬 IMAP search status: {status}", flush=True)
        if status == "OK":
            for num in messages[0].split():
                _, data = mail.fetch(num, "(RFC822)")
                msg = email.message_from_bytes(data[0][1])
                from_addr = email.utils.parseaddr(msg["From"])[1].lower().strip()
                replied.add(from_addr)
        mail.logout()
        print(f"✅ Found {len(replied)} new replies.", flush=True)
    except Exception as e:
        print(f"❌ IMAP Error while checking replies: {e}", flush=True)

    return replied

# === Google Sheets helpers ===
def hex_to_rgb(hex_color):
    return {
        "red": int(hex_color[0:2], 16) / 255,
        "green": int(hex_color[2:4], 16) / 255,
        "blue": int(hex_color[4:6], 16) / 255
    }

# --- UPDATED: Safe row color fetching ---
def get_all_row_colors(sheet_id, sheet_name, start_row=2, end_row=1000):
    print(f"🎨 Fetching row colors for rows {start_row} to {end_row}...", flush=True)
    try:
        range_ = f"{sheet_name}!A{start_row}:A{end_row}"
        result = sheets_api.spreadsheets().get(
            spreadsheetId=sheet_id,
            ranges=[range_],
            fields="sheets.data.rowData.values.effectiveFormat.backgroundColor"
        ).execute()
        row_colors = []
        row_data_list = result['sheets'][0]['data'][0].get('rowData', [])
        for row in row_data_list:
            if 'values' in row and len(row['values']) > 0:
                color = row['values'][0].get('effectiveFormat', {}).get('backgroundColor', {})
                rgb = (
                    int(color.get('red', 0) * 255),
                    int(color.get('green', 0) * 255),
                    int(color.get('blue', 0) * 255)
                )
            else:
                rgb = (255, 255, 255)
            row_colors.append(rgb)
        # Pad missing rows
        while len(row_colors) < (end_row - start_row + 1):
            row_colors.append((255, 255, 255))
        print(f"✅ Fetched {len(row_colors)} row colors.", flush=True)
        return row_colors
    except Exception as e:
        print(f"❌ Failed to fetch all row colors: {e}", flush=True)
        return [(255, 255, 255)] * (end_row - start_row + 1)

# === Batch Updates ===
def batch_update_cells(sheet_id, updates):
    print(f"🔄 Performing batch update on {len(updates)} cells...", flush=True)
    try:
        body = {"valueInputOption": "USER_ENTERED", "data": updates}
        sheets_api.spreadsheets().values().batchUpdate(
            spreadsheetId=sheet_id, body=body
        ).execute()
        print("✅ Batch update of cell values complete.", flush=True)
    except Exception as e:
        print(f"❌ Failed batch cell update: {e}", flush=True)

def batch_color_rows(spreadsheet_id, start_row_index_color_map, sheet_id):
    print(f"🔄 Coloring {len(start_row_index_color_map)} rows...", flush=True)
    requests = []
    for row_idx, hex_color in start_row_index_color_map.items():
        rgb = hex_to_rgb(hex_color)
        print(f"🎨 Coloring row {row_idx} with {hex_color} => RGB {rgb}", flush=True)
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": row_idx - 1, "endRowIndex": row_idx},
                "cell": {"userEnteredFormat": {"backgroundColor": rgb}},
                "fields": "userEnteredFormat.backgroundColor"
            }
        })
    try:
        response = sheets_api.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": requests}
        ).execute()
        print(f"✅ Batch row coloring done. Response: {json.dumps(response, indent=2)}", flush=True)
    except Exception as e:
        print(f"❌ Batch row coloring failed: {e}", flush=True)

# === Replies Processing ===
def process_replies():
    print("🔁 Processing replies...", flush=True)
    try:
        data = sheet.get_all_records()
        replied_emails = get_reply_emails()
        updates = []
        color_updates = {}
        row_colors = get_all_row_colors(sheet.spreadsheet.id, sheet.title, 2, len(data) + 1)
        for idx, row in enumerate(data, start=2):
            if not any(row.values()):
                print(f"⚠️ Row {idx} is empty, skipping...", flush=True)
                continue
            email_addr = row.get("Email", "").lower().strip()
            if not email_addr or row.get("Reply Status", "") == "Replied":
                print(f"⚠️ Row {idx}: Email missing or already Replied, skipping...", flush=True)
                continue
            rgb = row_colors[idx - 2]
            if rgb and rgb != (255, 255, 255):
                print(f"⚠️ Row {idx}: Already colored (RGB {rgb}), skipping...", flush=True)
                continue
            if email_addr in replied_emails:
                updates.append({"range": f"{sheet.title}!R{idx}", "values": [["Replied"]]})
                color_updates[idx] = "#FFFF00"
                print(f"✅ Row {idx}: Email {email_addr} marked as Replied.", flush=True)
        if updates:
            batch_update_cells(sheet.spreadsheet.id, updates)
        if color_updates:
            batch_color_rows(sheet.spreadsheet.id, color_updates, sheet._properties['sheetId'])
    except Exception as e:
        print(f"❌ Error in processing replies: {e}", flush=True)

# === Follow-ups Processing ===
def process_followups():
    print("🔁 Processing follow-up emails...", flush=True)
    try:
        data = sheet.get_all_records()
        today = datetime.today().strftime('%Y-%m-%d')
        updates = []
        color_updates = {}
        row_colors = get_all_row_colors(sheet.spreadsheet.id, sheet.title, 2, len(data) + 1)
        sent_tracker = set()

        for idx, row in enumerate(data, start=2):
            if not any(row.values()):
                print(f"⚠️ Row {idx} is empty, skipping...", flush=True)
                continue

            rgb = row_colors[idx - 2]
            if rgb and rgb != (255, 255, 255):
                print(f"⚠️ Row {idx}: Already colored (RGB {rgb}), skipping...", flush=True)
                continue

            email_addr = row.get("Email", "").lower().strip()
            if not email_addr or email_addr in sent_tracker:
                print(f"⚠️ Row {idx}: Email missing or already sent in this cycle, skipping...", flush=True)
                continue

            name = row.get("First_Name", "").strip()
            try:
                count = int(row.get("Follow-Up Count"))
                if count < 0:
                    count = 0
            except:
                count = 0

            last_date = row.get("Last Follow-Up Date", "")
            reply_status = row.get("Reply Status", "").strip()
            if reply_status in ["Replied", "No Reply After 4 Followups"]:
                print(f"⚠️ Row {idx}: Already replied or finished followups, skipping...", flush=True)
                continue

            if last_date:
                last_dt = datetime.strptime(last_date, "%Y-%m-%d")
                if (datetime.now() - last_dt).total_seconds() < 86400:
                    print(f"⚠️ Row {idx}: Last followup sent less than 24h ago, skipping...", flush=True)
                    continue

            if count >= 4:
                send_email(email_addr, "Should I Close Your File?", FINAL_EMAIL, name=name)
                updates.append({"range": f"{sheet.title}!R{idx}", "values": [["No Reply After 4 Followups"]]})
                color_updates[idx] = "#FF0000"
                sent_tracker.add(email_addr)
                print(f"❌ Row {idx}: Max follow-ups reached, final email sent.", flush=True)
                continue

            next_count = count
            try:
                followup_text = FOLLOWUP_EMAILS[next_count].replace("{%name%}", name)
                subject = FOLLOWUP_SUBJECTS[next_count]

                if next_count == 0:
                    show = row.get("Show", "").strip()
                    if not show:
                        print(f"⚠️ Row {idx}: Show missing for first followup, skipping...", flush=True)
                        continue
                    followup_text = followup_text.replace("{%show%}", show)
                    subject = subject.replace("{%show%}", show)
                elif next_count == 1:
                    url = row.get("Pitch Deck URL", "").strip()
                    if not url:
                        print(f"⚠️ Row {idx}: Pitch Deck URL missing, skipping second followup...", flush=True)
                        continue
                    followup_text = followup_text.replace("{%pitch_deck_url%}", url)

                send_email(email_addr, subject, followup_text, name=name)
                sent_tracker.add(email_addr)
                print(f"✅ Row {idx}: Sent followup {next_count+1} to {email_addr}", flush=True)

                updates.extend([
                    {"range": f"{sheet.title}!P{idx}", "values": [[str(next_count + 1)]]},
                    {"range": f"{sheet.title}!Q{idx}", "values": [[today]]},
                    {"range": f"{sheet.title}!R{idx}", "values": [["Pending"]]}
                ])

            except Exception as e:
                print(f"❌ Failed to prepare/send follow-up email to {email_addr}: {e}", flush=True)
                continue

            if (idx - 1) % 3 == 0:
                print("⏱ Sleeping 3 seconds between emails...", flush=True)
                time.sleep(3)

        if updates:
            batch_update_cells(sheet.spreadsheet.id, updates)
        if color_updates:
            batch_color_rows(sheet.spreadsheet.id, color_updates, sheet._properties['sheetId'])

    except Exception as e:
        print(f"❌ Error in processing followups: {e}", flush=True)

# === MAIN LOOP ===
if __name__ == "__main__":
    print("🚀 Sales follow-up automation started...", flush=True)
    next_followup_check = time.time()
    while True:
        try:
            print("\n--- Checking for replies ---", flush=True)
            process_replies()

            current_time = time.time()
            if current_time >= next_followup_check:
                print("\n--- Sending follow-up emails ---", flush=True)
                process_followups()
                next_followup_check = current_time + 86400  # 24h

        except Exception:
            print("❌ Fatal error:", flush=True)
            traceback.print_exc()

        print("⏱ Sleeping 30 seconds before next cycle...", flush=True)
        time.sleep(30)
