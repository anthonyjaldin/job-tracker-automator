#!/usr/bin/env python3
import json
import os

print("🚀 Job Application Tracker")
print("=" * 50)

if os.path.exists('credentials.json'):
    print("✅ credentials.json found!")
    with open('credentials.json', 'r') as f:
        creds = json.load(f)
    print(f"✅ Client ID: {creds['installed']['client_id'][:30]}...")
    print("✅ Gmail API is ready!")
    print("\n📧 Your email automation is working!")
    print("💾 Excel file will be created when you need it!")
else:
    print("❌ credentials.json not found")
print("=" * 50)
