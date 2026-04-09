import os
import json
import smtplib
import ssl
from email.message import EmailMessage
import gspread
from google.oauth2.service_account import Credentials
from jinja2 import Template
import base64

# Configuration
SPREADSHEET_ID = '11vFXr-7xKguLY_sa9Gi6Ne1AgA78MiY89RiVON_jrQk'
SMTP_HOST = 'secure.emailsrvr.com'
SMTP_PORT = 465
SENDER_EMAIL = 'mayur.kambli@artworkservicesusa.com'

# Recipients
RECIPIENTS_TO = ['mayur.online9@gmail.com']
RECIPIENTS_CC = ['mayur.kambli@artworkservicesusa.com']
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
    rows = all_data[1:]
    
    emails_received = 0
    emails_completed = 0
    total_completed_items = 0
    pending_count = 0
    detailed_rows = []

    for row in rows:
        if len(row) < 15: # Pad row if shorter than index O
            row = row + [""] * (15 - len(row))
        
        row_date = str(row[0]).strip()
        done_date = str(row[14]).strip()
        total_items = str(row[7]).strip()
        
        # 1. Emails Received: Date matches target_date
        if row_date == target_date:
            emails_received += 1
            
        # 2. Emails Completed: Done date matches target_date
        # 3. Total Items: Sum where Done date matches target_date
        if done_date == target_date:
            emails_completed += 1
            try:
                total_completed_items += int(total_items) if total_items else 0
            except:
                pass
        
        # 4. Pending: Done date is "Pending" or empty (but has a Date in col A)
        if (done_date.lower() == 'pending' or not done_date) and row_date:
            pending_count += 1
            
        # 5. Detailed Rows: Include if Done date == target_date, OR Done date is blank/Pending
        #    (blank Done Date = always include, regardless of received date)
        is_pending = not done_date or done_date.lower() == 'pending'
        if (done_date == target_date or is_pending) and row_date:
            # Format: [Date, Email From, Email Subject, Total items, Done date]
            # Count should be blank for Pending rows
            detailed_rows.append([
                row[0], 
                row[5], 
                row[6], 
                row[7] if not is_pending else "", 
                row[14] if row[14].strip() else "Pending"
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
            <p>Hi Rupesh,</p>
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
        <p>Hi Rupesh,</p>
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
