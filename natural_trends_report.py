import os
import json
import smtplib
import ssl
from email.message import EmailMessage
import gspread
from google.oauth2.service_account import Credentials
from jinja2 import Template
import base64
from datetime import datetime

# Configuration
SPREADSHEET_ID = '11vFXr-7xKguLY_sa9Gi6Ne1AgA78MiY89RiVON_jrQk'
SMTP_HOST = 'secure.emailsrvr.com'
SMTP_PORT = 465
SENDER_EMAIL = 'mayur.kambli@artworkservicesusa.com'

# Recipients
RECIPIENTS_TO = ['sharedart@naturaltrends.com']
RECIPIENTS_CC = ['nikki@naturaltrends.com', 'joel@naturaltrends.com', 'victoria@naturaltrends.com', 'rupesh.pardeshi@artworkservicesusa.com', 'ashok.sharma@artworkservicesusa.com']
RECIPIENTS_BCC = ['mayur.kambli@artworkservicesusa.com']

def get_google_client():
    creds_raw = os.environ.get('GOOGLE_CREDENTIALS_JSON', '').strip()
    if not creds_raw:
        raise ValueError("GOOGLE_CREDENTIALS_JSON environment variable not set")
    
    creds_dict = None
    # 1. Try Base64 decoding
    try:
        decoded = base64.b64decode(creds_raw).decode('utf-8')
        if decoded.strip().startswith('{'):
            creds_dict = json.loads(decoded)
            print("Successfully loaded credentials via Base64.")
    except Exception:
        pass

    # 2. Fallback to raw JSON (with cleanup for literal newlines)
    if not creds_dict:
        try:
            cleaned_raw = creds_raw.replace('\r\n', '\\n').replace('\n', '\\n')
            creds_dict = json.loads(cleaned_raw)
            print("Successfully loaded credentials via Raw JSON (auto-fixed).")
        except Exception:
            try:
                creds_dict = json.loads(creds_raw)
                print("Successfully loaded credentials via Raw JSON.")
            except json.JSONDecodeError as e:
                raise ValueError(f"CRITICAL: GOOGLE_CREDENTIALS_JSON is malformed. Error: {e}")

    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

def fetch_data():
    gc = get_google_client()
    sh = gc.open_by_key(SPREADSHEET_ID)
    
    # 1. Get Target Date from D.count!Z2
    ws_count = sh.worksheet("D.count")
    target_date = ws_count.acell('Z2').value
    print(f"Target Date (D.count!Z2): {target_date}")
    
    # 2. Get Data from Natural Trends
    ws_trends = sh.worksheet("Natural Trends")
    # Fetch all data (skipping the header row 1 if it's always the same)
    # Looking at the image, row 1, 5, 9 etc. are headers/red bars.
    # Actually, let's get all and filter programmatically.
    all_data = ws_trends.get_all_values()
    if not all_data:
        return target_date, 0, 0, 0, 0, []

    # Map headers to column indices based on image:
    # A(0): Date, D(3): Email From, E(4): Email subject, F(5): Total items, M(12): Done date
    headers = all_data[0]
    def normalize_date(d):
        d = d.strip()
        if not d: return ""
        # Handle formats like 1-May-26 and 01-May-26
        parts = d.split('-')
        if len(parts) == 3:
            day = parts[0].lstrip('0')
            return f"{day}-{parts[1]}-{parts[2]}"
        return d

    target_date_norm = normalize_date(target_date)

    rows = all_data[1:]
    
    emails_received = 0
    emails_completed = 0
    total_completed_items = 0
    pending_count = 0
    detailed_rows = []

    today = datetime.now()
    today_str = f"{today.day}-{today.strftime('%b-%y')}"

    for row in rows:
        if len(row) < 16: # Pad row if shorter than index P (15)
            row = row + [""] * (16 - len(row))
        
        row_date_raw = str(row[0]).strip()
        row_date = normalize_date(row_date_raw)
        
        if row_date == today_str:
            continue
            
        email_subject = str(row[6]).strip()
        
        done_date_raw = str(row[15]).strip() # Column P
        done_date = normalize_date(done_date_raw)
        
        total_items = str(row[7]).strip() # Column H
        
        # 1. Emails Received: Date matches target_date
        if row_date == target_date_norm:
            emails_received += 1
            
        # 2. Emails Completed: Done date matches target_date
        if done_date == target_date_norm:
            emails_completed += 1
            try:
                # Clean total_items string
                val = "".join(filter(str.isdigit, total_items))
                total_completed_items += int(val) if val else 0
            except:
                pass
        
        # 3. Pending: Done date is "Pending" or empty (but has a Date in col A)
        is_pending = not done_date or done_date.lower() == 'pending'
        if is_pending and row_date and email_subject:
            pending_count += 1
            
        # 4. Detailed Rows: ONLY include items for the target date
        if (row_date == target_date_norm or done_date == target_date_norm):
            detailed_rows.append([
                row_date_raw, 
                row[5], # Email From
                row[6], # Email Subject
                row[7] if not is_pending else "", 
                done_date_raw if done_date_raw.strip() else "Pending"
            ])

    return target_date, emails_received, emails_completed, total_completed_items, pending_count, detailed_rows

def format_html(target_date, emails_received, emails_completed, total_completed_items, pending_count, detailed_rows):
    # Template Case 2: No orders received for the target date
    if emails_received == 0 and emails_completed == 0:
        template_str = """
        <html>
        <head>
        <style>
            body { font-family: Calibri, sans-serif; font-size: 10pt; line-height: 1.2; }
            table { border-collapse: collapse; border: 1px solid #000000; margin-top: 10px; }
            td { border: 1px solid #000000; padding: 2px 6px; font-size: 10pt; }
            .header-cell { background-color: #ffffff; font-weight: bold; }
        </style>
        </head>
        <body>
            <p>Hi team,</p>
            <p>Please see below summary.</p>
            <table>
                <tr><td class="header-cell">Natural Trends</td><td></td></tr>
                <tr><td class="header-cell">Date</td><td class="header-cell">Emails received</td></tr>
                <tr><td>{{ target_date }}</td><td>No orders received</td></tr>
            </table>
            <br>
            <p>Thanks and Regards,<br>Mayur</p>
        </body>
        </html>
        """
        template = Template(template_str)
        return template.render(target_date=target_date)

    # Template Case 1: Data found
    template_str = """
    <html>
    <head>
    <style>
        body { font-family: Calibri, sans-serif; font-size: 10pt; line-height: 1.2; }
        table { border-collapse: collapse; border: 1px solid #000000; margin-top: 10px; width: auto; min-width: 400px; }
        td { border: 1px solid #000000; padding: 2px 6px; font-size: 10pt; }
        .header-cell { font-weight: bold; }
        .detail-table { width: 100%; max-width: 800px; }
    </style>
    </head>
    <body>
        <p>Hi team,</p>
        <p>Please see below mentioned summary report for your reference.</p>
        
        <!-- Summary Table -->
        <table>
            <tr><td colspan="4" class="header-cell">Natural Trends</td></tr>
            <tr class="header-cell">
                <td>Date</td><td>Emails received</td><td>Emails completed</td><td>Total virtual/proofs completed</td>
            </tr>
            <tr>
                <td>{{ target_date }}</td><td>{{ emails_received }}</td><td>{{ emails_completed }}</td><td>{{ total_completed_items }}</td>
            </tr>
            <tr><td>&nbsp;</td><td></td><td></td><td></td></tr>
            <tr><td class="header-cell">Pending</td><td>{{ pending_count }}</td><td></td><td></td></tr>
        </table>

        <br>

        <!-- Detailed Table -->
        <table class="detail-table">
            <tr><td colspan="5" class="header-cell" style="text-align: center;">Natural Trends</td></tr>
            <tr class="header-cell">
                <td>Date</td><td>Emails from</td><td>Email subject</td><td>Count</td><td>Done date</td>
            </tr>
            {% for row in detailed_rows %}
            <tr>
                <td>{{ row[0] }}</td><td>{{ row[1] }}</td><td>{{ row[2] }}</td><td>{{ row[3] }}</td><td>{{ row[4] }}</td>
            </tr>
            {% endfor %}
        </table>
        
        <br>
        <p>Thanks and Regards,<br>Mayur</p>
    </body>
    </html>
    """
    template = Template(template_str)
    return template.render(
        target_date=target_date,
        emails_received=emails_received,
        emails_completed=emails_completed,
        total_completed_items=total_completed_items,
        pending_count=pending_count,
        detailed_rows=detailed_rows
    )

def send_email(subject, html_content):
    password = os.environ.get('SMTP_PASSWORD')
    if not password:
        raise ValueError("SMTP_PASSWORD environment variable not set")
    
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = SENDER_EMAIL
    msg['To'] = ", ".join(RECIPIENTS_TO)
    msg['Cc'] = ", ".join(RECIPIENTS_CC)
    msg['Bcc'] = ", ".join(RECIPIENTS_BCC)
    
    msg.set_content("Please enable HTML to view this report.")
    msg.add_alternative(html_content, subtype='html')
    
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
        server.login(SENDER_EMAIL, password)
        server.send_message(msg)

def main():
    try:
        print("Fetching data from Google Sheets...")
        target_date, emails_received, emails_completed, total_items, pending, detailed = fetch_data()
        
        print(f"Generating report for {target_date}...")
        html_content = format_html(target_date, emails_received, emails_completed, total_items, pending, detailed)
        
        subject = f"Natural Trends Summary: {target_date}"
        print(f"Sending email: {subject}")
        send_email(subject, html_content)
        
        print("Success!")
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        exit(1)

if __name__ == "__main__":
    main()
