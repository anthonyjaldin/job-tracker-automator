#!/usr/bin/env python3
"""
Real Job Application Tracker - Scans Gmail and updates Excel automatically
"""
import json
import base64
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
except ImportError:
    print("Installing required packages...")
    os.system("pip install google-auth-oauthlib google-auth-httplib2 google-api-python-client --break-system-packages")
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
COMPANIES = ['Amazon', 'Microsoft', 'Google', 'Apple', 'Meta', 'Tesla', 'IBM', 'Oracle', 'Cisco', 'Intel', 'Capital One', 'JPMorgan', 'Goldman Sachs', 'Deloitte', 'Accenture']

def authenticate():
    """Authenticate with Gmail API"""
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds

def scan_sent_emails(service):
    """Scan sent emails for job applications"""
    try:
        results = service.users().messages().list(userId='me', q='in:sent', maxResults=50).execute()
        messages = results.get('messages', [])
        
        applications = []
        
        for msg in messages:
            msg_data = service.users().messages().get(userId='me', id=msg['id']).execute()
            headers = msg_data['payload']['headers']
            
            subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '')
            date = next((h['value'] for h in headers if h['name'] == 'Date'), '')
            
            # Check if email contains company names or job keywords
            for company in COMPANIES:
                if company.lower() in subject.lower():
                    applications.append({
                        'date': date[:10],
                        'company': company,
                        'position': subject,
                        'status': 'Applied',
                        'updated': datetime.now().strftime('%m/%d/%Y')
                    })
                    break
        
        return applications
    except Exception as e:
        print(f"Error scanning emails: {e}")
        return []

def update_excel(applications):
    """Update Excel file with applications"""
    if not os.path.exists('job_applications_tracker.xlsx'):
        print("Excel file not found!")
        return
    
    wb = load_workbook('job_applications_tracker.xlsx')
    ws = wb.active
    
    # Clear existing data (keep headers)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.value = None
    
    # Add new applications
    row_num = 2
    status_colors = {
        'Applied': 'D4EDDA',
        'Pending': 'FFF3CD',
        'Interview': 'B8DAFF',
        'Rejected': 'F8D7DA',
        'Offer': 'D1E7DD'
    }
    
    for app in applications:
        ws[f'A{row_num}'] = app['date']
        ws[f'B{row_num}'] = app['company']
        ws[f'C{row_num}'] = app['position']
        ws[f'D{row_num}'] = app['status']
        ws[f'E{row_num}'] = app['updated']
        
        # Color code status
        status_fill = PatternFill(start_color=status_colors.get(app['status'], 'FFFFFF'),
                                 end_color=status_colors.get(app['status'], 'FFFFFF'),
                                 fill_type="solid")
        ws[f'D{row_num}'].fill = status_fill
        
        row_num += 1
    
    wb.save('job_applications_tracker.xlsx')
    print(f"✅ Updated Excel with {len(applications)} applications")

def main():
    print("🚀 Job Application Tracker - Live Scanner")
    print("=" * 60)
    
    try:
        print("🔐 Authenticating with Gmail...")
        creds = authenticate()
        print("✅ Gmail authenticated!")
        
        print("📧 Scanning sent emails...")
        service = build('gmail', 'v1', credentials=creds)
        applications = scan_sent_emails(service)
        
        print(f"✅ Found {len(applications)} job applications")
        
        print("📊 Updating Excel...")
        update_excel(applications)
        
        print("=" * 60)
        print("✅ Scanner complete!")
        print(f"📁 Check: job_applications_tracker.xlsx")
        
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == '__main__':
    main()
