#!/usr/bin/env python3
"""
Quick Email Scanner - Manual scan without offline access
"""
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def manual_add_application():
    """Manually add an application to Excel"""
    print("🚀 Manual Job Application Tracker")
    print("=" * 60)
    
    if not os.path.exists('job_applications_tracker.xlsx'):
        print("❌ Excel file not found!")
        return
    
    print("\nAdd your job application:")
    date = input("📅 Application Date (MM/DD/YYYY): ").strip()
    company = input("🏢 Company Name: ").strip()
    position = input("💼 Position: ").strip()
    status = input("📊 Status (Applied/Pending/Interview/Rejected/Offer): ").strip()
    notes = input("📝 Notes (optional): ").strip()
    
    wb = load_workbook('job_applications_tracker.xlsx')
    ws = wb.active
    
    # Find next empty row
    row_num = 2
    while ws[f'A{row_num}'].value is not None:
        row_num += 1
    
    # Add data
    ws[f'A{row_num}'] = date
    ws[f'B{row_num}'] = company
    ws[f'C{row_num}'] = position
    ws[f'D{row_num}'] = status
    ws[f'E{row_num}'] = datetime.now().strftime('%m/%d/%Y')
    ws[f'F{row_num}'] = notes
    
    # Color code
    status_colors = {
        'Applied': 'D4EDDA',
        'Pending': 'FFF3CD',
        'Interview': 'B8DAFF',
        'Rejected': 'F8D7DA',
        'Offer': 'D1E7DD'
    }
    
    color = status_colors.get(status, 'FFFFFF')
    ws[f'D{row_num}'].fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    
    wb.save('job_applications_tracker.xlsx')
    
    print("\n" + "=" * 60)
    print("✅ Application added to Excel!")
    print(f"📁 Updated: job_applications_tracker.xlsx")

if __name__ == '__main__':
    manual_add_application()
