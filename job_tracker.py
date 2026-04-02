"""
Job Application Tracker Automator
Processes recruitment emails and organizes job opportunities into an Excel spreadsheet
"""

import os
import base64
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.auth.oauthlib.flow import InstalledAppFlow
from google_auth_httplib2 import AuthorizedHttp
import httplib2
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from email.mime.text import MIMEText
import re

# Gmail API scope
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.json'
OUTPUT_FILE = 'job_applications.xlsx'

# Keywords to identify job/recruitment emails
JOB_KEYWORDS = [
    'apply', 'job', 'position', 'opportunity', 'recruitment',
    'hiring', 'career', 'internship', 'role', 'candidate',
    'application status', 'we\'re interested', 'interview'
]


def authenticate_gmail():
    """Authenticate with Gmail API"""
    creds = None
    
    # Token.json stores refresh token
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    
    # If credentials don't exist, create them
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save credentials for next run
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    
    return creds


def get_emails(creds, query='from:*', max_results=10):
    """Fetch emails from Gmail matching query"""
    service = None
    try:
        import googleapiclient.discovery
        service = googleapiclient.discovery.build('gmail', 'v1', credentials=creds)
    except:
        print("Error: Install google-auth-httplib2 and google-auth-oauthlib")
        return []
    
    try:
        results = service.users().messages().list(userId='me', q=query, maxResults=max_results).execute()
        messages = results.get('messages', [])
        return service, messages
    except Exception as e:
        print(f"Error fetching emails: {e}")
        return None, []


def extract_email_details(service, message_id):
    """Extract sender, subject, date, and body from email"""
    try:
        message = service.users().messages().get(userId='me', id=message_id, format='full').execute()
        headers = message['payload']['headers']
        
        # Extract header data
        sender = next((h['value'] for h in headers if h['name'] == 'From'), 'Unknown')
        subject = next((h['value'] for h in headers if h['name'] == 'Subject'), 'No Subject')
        date_str = next((h['value'] for h in headers if h['name'] == 'Date'), datetime.now().isoformat())
        
        # Try to extract body
        body = extract_body(message['payload'])
        
        return {
            'sender': sender,
            'subject': subject,
            'date': date_str,
            'body': body
        }
    except Exception as e:
        print(f"Error extracting email: {e}")
        return None


def extract_body(payload):
    """Recursively extract email body text"""
    try:
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'text/plain':
                    data = part['body'].get('data', '')
                    if data:
                        return base64.urlsafe_b64decode(data).decode('utf-8')
        else:
            data = payload['body'].get('data', '')
            if data:
                return base64.urlsafe_b64decode(data).decode('utf-8')
    except:
        pass
    return ''


def is_job_email(subject, body):
    """Determine if email is job-related"""
    text = f"{subject} {body}".lower()
    return any(keyword in text for keyword in JOB_KEYWORDS)


def extract_company_name(sender):
    """Extract company name from email sender"""
    # Try to extract from email domain
    if '@' in sender:
        domain = sender.split('@')[1].split('>')[0]
        return domain.replace('.com', '').replace('.org', '').title()
    return sender.split('<')[0].strip() if '<' in sender else sender


def parse_email_date(date_str):
    """Parse email date string to datetime"""
    try:
        # Try common email date formats
        from email.utils import parsedate_to_datetime
        return parsedate_to_datetime(date_str)
    except:
        return datetime.now()


def create_excel_sheet(jobs):
    """Create organized Excel spreadsheet from job data"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Job Applications"
    
    # Define styles
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Column headers
    headers = ['Date', 'Company', 'Position/Subject', 'From', 'Status', 'Notes']
    ws.append(headers)
    
    # Style headers
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    
    # Set column widths
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 20
    
    # Add data rows (sorted by date)
    for job in sorted(jobs, key=lambda x: x['date'], reverse=True):
        ws.append([
            job['date'].strftime('%m/%d/%Y'),
            job['company'],
            job['subject'][:50],  # Truncate long subjects
            job['sender'][:30],
            'Applied',
            ''  # Empty notes column for user to fill in
        ])
    
    # Style data rows
    for row in ws.iter_rows(min_row=2, max_row=len(jobs)+1):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    # Freeze header row
    ws.freeze_panes = 'A2'
    
    # Save file
    wb.save(OUTPUT_FILE)
    print(f"✅ Excel file created: {OUTPUT_FILE}")
    print(f"📊 Total job applications tracked: {len(jobs)}")


def main():
    """Main execution"""
    print("🚀 Job Application Tracker - Email Automator")
    print("=" * 50)
    
    # Check if credentials exist
    if not os.path.exists(CREDENTIALS_FILE):
        print("⚠️  Please download credentials from Google Cloud Console")
        print("   1. Go to https://console.cloud.google.com/")
        print("   2. Create OAuth 2.0 credentials (Desktop App)")
        print("   3. Download as JSON and save as 'credentials.json'")
        return
    
    # Authenticate
    print("🔐 Authenticating with Gmail...")
    creds = authenticate_gmail()
    
    # Fetch emails
    print("📧 Fetching emails...")
    service, messages = get_emails(creds, query='is:unread', max_results=20)
    
    if not service:
        print("❌ Failed to connect to Gmail API")
        return
    
    # Process emails
    print(f"📬 Processing {len(messages)} emails...")
    jobs = []
    
    for msg in messages:
        details = extract_email_details(service, msg['id'])
        if not details:
            continue
        
        # Check if email is job-related
        if is_job_email(details['subject'], details['body']):
            job_entry = {
                'date': parse_email_date(details['date']),
                'company': extract_company_name(details['sender']),
                'subject': details['subject'],
                'sender': details['sender'],
            }
            jobs.append(job_entry)
            print(f"  ✓ {job_entry['company']} - {job_entry['subject'][:40]}")
    
    if jobs:
        # Create Excel spreadsheet
        print(f"\n📊 Creating Excel spreadsheet with {len(jobs)} job opportunities...")
        create_excel_sheet(jobs)
        print("\n✅ Job tracker automation complete!")
    else:
        print("\n⚠️  No job-related emails found in inbox")


if __name__ == '__main__':
    main()
