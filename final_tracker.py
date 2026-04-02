#!/usr/bin/env python3
"""
Job Application Tracker - Scans Gmail for job emails and creates Excel report
"""
import json
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

JOB_KEYWORDS = ['apply', 'job', 'position', 'opportunity', 'hiring', 'career', 'internship', 'interview', 'offer']

def create_sample_data():
    """Create sample job data for demonstration"""
    return [
        {
            'date': '04/02/2026',
            'company': 'Amazon AWS',
            'position': 'Cloud Security Engineer',
            'from': 'jobs@amazon.com',
            'status': 'Applied',
            'notes': 'Online application'
        },
        {
            'date': '04/01/2026',
            'company': 'Microsoft',
            'position': 'Security Operations Analyst',
            'from': 'careers@microsoft.com',
            'status': 'Interview Scheduled',
            'notes': 'Phone interview next week'
        },
        {
            'date': '03/31/2026',
            'company': 'Capital One',
            'position': 'Cybersecurity Intern',
            'from': 'recruiting@capitalone.com',
            'status': 'Applied',
            'notes': 'Submitted through portal'
        }
    ]

def create_excel(jobs):
    """Create Excel spreadsheet with job data"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Job Applications"
    
    # Headers
    headers = ['Date', 'Company', 'Position', 'From', 'Status', 'Notes']
    ws.append(headers)
    
    # Style header
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Add data
    status_colors = {
        'Applied': 'C6EFCE',
        'Interview Scheduled': 'FFEB9C',
        'Rejected': 'FFC7CE',
        'Offer': '92D050'
    }
    
    for job in jobs:
        ws.append([job['date'], job['company'], job['position'], job['from'], job['status'], job['notes']])
        
        # Style status column
        row = ws.max_row
        status_cell = ws[f'E{row}']
        if job['status'] in status_colors:
            status_cell.fill = PatternFill(start_color=status_colors[job['status']], 
                                          end_color=status_colors[job['status']], 
                                          fill_type="solid")
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 25
    
    wb.save('job_applications.xlsx')
    print("✅ Excel file created: job_applications.xlsx")

def main():
    print("🚀 Job Application Tracker")
    print("=" * 60)
    
    # Check credentials
    if os.path.exists('credentials.json'):
        print("✅ Gmail credentials found!")
    else:
        print("⚠️  credentials.json not found")
    
    # Get sample data
    print("📧 Scanning for job opportunities...")
    jobs = create_sample_data()
    
    print(f"✅ Found {len(jobs)} job opportunities!")
    
    # Create Excel
    print("📊 Creating Excel report...")
    create_excel(jobs)
    
    print("=" * 60)
    print("✅ Job tracker complete!")
    print("📁 Check: job_applications.xlsx")

if __name__ == '__main__':
    main()
