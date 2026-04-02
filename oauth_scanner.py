#!/usr/bin/env python3
"""
Job Application Scanner with proper OAuth for WSL
"""
import os
import webbrowser
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

COMPANIES = {
    'amazon': 'Amazon',
    'microsoft': 'Microsoft',
    'google': 'Google',
    'capital one': 'Capital One',
    'jpmorgan': 'JPMorgan',
    'apple': 'Apple',
    'meta': 'Meta',
    'tesla': 'Tesla',
    'ibm': 'IBM',
    'oracle': 'Oracle',
    'cisco': 'Cisco',
    'intel': 'Intel',
    'deloitte': 'Deloitte',
    'accenture': 'Accenture'
}

def authenticate():
    """Authenticate without offline access"""
    creds = None
    
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json',
                SCOPES,
                redirect_uri='http://localhost:8080'
            )
            # Don't use offline - just get auth
            auth_url, _ = flow.authorization_url(access_type='online')
            print(f"\n🔗 Please open this URL in your browser:\n{auth_url}\n")
            webbrowser.open(auth_url)
            
            auth_code = input("📋 Paste the authorization code here: ").strip()
            creds = flow.fetch_token(code=auth_code)
        
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds

def scan_emails(creds):
    """Scan Gmail for job applications"""
    print("\n📧 Scanning your sent emails...")
    
    service = build('gmail', 'v1', credentials=creds)
    
    try:
        results = service.users().messages().list(
            userId='me',
            q='in:sent',
            maxResults=100
        ).execute()
        
        messages = results.get('messages', [])
        applications = []
        
        for msg in messages:
            msg_data = service.users().messages().get(userId='me', id=msg['id']).execute()
            headers = msg_data['payload'].get('headers', [])
            
            subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '').lower()
            date_str = next((h['value'] for h in headers if h['name'] == 'Date'), '')
            
            # Parse date
            date_display = date_str[:10] if date_str else datetime.now().strftime('%m/%d/%Y')
            
            # Find company
            for company_key, company_name in COMPANIES.items():
                if company_key in subject:
                    applications.append({
                        'date': date_display,
                        'company': company_name,
                        'subject': subject.title()[:50],
                        'status': 'Applied'
                    })
                    print(f"  ✓ {company_name}")
                    break
        
        return applications
    
    except Exception as e:
        print(f"❌ Error: {e}")
        return []

def update_excel(applications):
    """Update Excel file"""
    if not os.path.exists('job_applications_tracker.xlsx'):
        print("❌ Excel file not found!")
        return
    
    print(f"\n📊 Adding {len(applications)} applications to Excel...")
    
    wb = load_workbook('job_applications_tracker.xlsx')
    ws = wb.active
    
    status_colors = {
        'Applied': 'D4EDDA',
        'Pending': 'FFF3CD',
        'Interview': 'B8DAFF',
        'Rejected': 'F8D7DA',
        'Offer': 'D1E7DD'
    }
    
    row_num = 2
    for app in applications:
        ws[f'A{row_num}'] = app['date']
        ws[f'B{row_num}'] = app['company']
        ws[f'C{row_num}'] = app['subject']
        ws[f'D{row_num}'] = app['status']
        ws[f'E{row_num}'] = datetime.now().strftime('%m/%d/%Y')
        
        color = status_colors.get(app['status'], 'FFFFFF')
        ws[f'D{row_num}'].fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        
        row_num += 1
    
    wb.save('job_applications_tracker.xlsx')
    print("✅ Excel updated!")

def main():
    print("🚀 Job Application Scanner - OAuth")
    print("=" * 60)
    
    try:
        print("🔐 Authenticating with Gmail...")
        creds = authenticate()
        print("✅ Gmail authenticated!")
        
        applications = scan_emails(creds)
        
        if applications:
            update_excel(applications)
            print("\n" + "=" * 60)
            print("✅ Done! Check: job_applications_tracker.xlsx")
        else:
            print("⚠️  No job applications found")
        
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == '__main__':
    main()
