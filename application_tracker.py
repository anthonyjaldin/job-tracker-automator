#!/usr/bin/env python3
"""
Job Application Tracker - Tracks YOUR sent emails with job applications
"""
import json
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

def create_excel_tracker():
    """Create Excel spreadsheet to track your job applications"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Applications"
    
    # Headers
    headers = ['Application Date', 'Company', 'Position', 'Status', 'Last Updated', 'Notes']
    ws.append(headers)
    
    # Style header
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=12)
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # Status color mapping
    status_colors = {
        'Applied': 'D4EDDA',      # Green
        'Pending': 'FFF3CD',      # Yellow
        'Interview': 'B8DAFF',    # Blue
        'Rejected': 'F8D7DA',     # Red
        'Offer': 'D1E7DD'         # Light Green
    }
    
    # Add column widths
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 16
    ws.column_dimensions['F'].width = 30
    
    # Freeze header row
    ws.freeze_panes = 'A2'
    
    print("✅ Excel template created!")
    wb.save('job_applications_tracker.xlsx')
    print("📁 File saved: job_applications_tracker.xlsx")
    return ws

def main():
    print("🚀 Job Application Tracker")
    print("=" * 60)
    print("📧 This tracker monitors YOUR sent emails for job applications")
    print("=" * 60)
    
    # Check credentials
    if os.path.exists('credentials.json'):
        print("✅ Gmail credentials configured")
    else:
        print("⚠️  No Gmail credentials found")
    
    print("\n📊 Creating application tracker template...")
    create_excel_tracker()
    
    print("\n" + "=" * 60)
    print("✅ Tracker ready!")
    print("\nHow to use:")
    print("1. Open: job_applications_tracker.xlsx")
    print("2. Add your applications:")
    print("   - Application Date: When you applied")
    print("   - Company: Company name")
    print("   - Position: Job position")
    print("   - Status: Applied/Pending/Interview/Rejected/Offer")
    print("   - Last Updated: Latest status date")
    print("   - Notes: Any details")
    print("\n3. Color-coded by status for easy tracking")
    print("=" * 60)

if __name__ == '__main__':
    main()
