# Job Application Tracker Automator

A Python automation tool that scans your Gmail inbox for recruitment emails and automatically organizes job opportunities into a professional Excel spreadsheet. Built to help job seekers efficiently track applications, dates, and follow-ups.

## Features

✨ **Automated Email Scanning** - Monitors Gmail for job-related emails  
📊 **Smart Organization** - Extracts company, position, sender, and dates  
📅 **Date Sorting** - Applications organized chronologically (newest first)  
🎨 **Professional Formatting** - Color-coded status tracking and styled Excel output  
🔐 **Secure OAuth 2.0** - Uses Google's secure authentication (no password storage)  
⚡ **One-Click Operation** - Run once and get organized tracking spreadsheet  

## Prerequisites

- Python 3.7 or higher
- Google Gmail account
- GitHub account (for cloning this repo)

## Installation & Setup

### Step 1: Clone the Repository

```bash
git clone https://github.com/YOUR-USERNAME/job-tracker-automator.git
cd job-tracker-automator
```

### Step 2: Install Python Dependencies

```bash
pip install google-auth-oauthlib google-auth-httplib2 google-api-python-client openpyxl
```

Or use the requirements file:
```bash
pip install -r requirements.txt
```

### Step 3: Set Up Google Cloud Project & Gmail API

**3a. Create a Google Cloud Project**

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Click "Select a Project" at the top → Click "NEW PROJECT"
3. Name it: `Job Tracker Automator`
4. Click "CREATE"
5. Wait for the project to be created (this takes ~1 minute)

**3b. Enable Gmail API**

1. In the Cloud Console, search for "Gmail API" in the search bar at top
2. Click on "Gmail API"
3. Click the blue "ENABLE" button
4. This enables the API for your project

**3c. Create OAuth 2.0 Credentials**

1. In the Cloud Console, go to **APIs & Services** → **Credentials** (left sidebar)
2. Click **"+ CREATE CREDENTIALS"** button at the top
3. Select **"OAuth client ID"**
4. If prompted, click "Configure Consent Screen" first:
   - Select "External" user type
   - Click "CREATE"
   - Fill in basic info (App name: "Job Tracker", User support email: your email)
   - Under "Scopes," click "ADD OR REMOVE SCOPES"
   - Search for and add: `https://www.googleapis.com/auth/gmail.readonly`
   - Click "UPDATE"
   - Click "SAVE AND CONTINUE"
   - Go back to "Credentials"

5. Now click **"+ CREATE CREDENTIALS"** → **"OAuth client ID"** again
6. Select **"Desktop application"** as the application type
7. Click "CREATE"
8. Click the download icon (⬇️) next to your OAuth client
9. Save it as **`credentials.json`** in your project folder

**3d. Verify Your Setup**

Your project folder should now contain:
```
job-tracker-automator/
├── job_tracker.py
├── credentials.json          ← Downloaded from Google Cloud
├── requirements.txt
├── README.md
└── .gitignore              ← Protects credentials.json
```

### Step 4: Run the Automator

```bash
python job_tracker.py
```

**On first run:**
- A browser window will automatically open asking for permission
- Click "Allow" to grant access to your Gmail
- The app will save a `token.json` file for future runs (no re-authentication needed)

**Output:**
- Creates `job_applications.xlsx` in your project folder
- Shows summary in terminal with company names and positions found

## Usage

### Basic Run
```bash
python job_tracker.py
```

### Customize Search Queries

Edit `job_tracker.py` to change the email search query:

```python
# Search all unread emails (default)
service, messages = get_emails(creds, query='is:unread', max_results=20)

# Or search by date range
service, messages = get_emails(creds, query='after:2026/03/01', max_results=50)

# Or search specific sender
service, messages = get_emails(creds, query='from:jobs@amazon.com', max_results=20)
```

### Advanced: Modify Job Keywords

Edit the `JOB_KEYWORDS` list in `job_tracker.py` to catch more/fewer emails:

```python
JOB_KEYWORDS = [
    'apply', 'job', 'position', 'opportunity', 'recruitment',
    'hiring', 'career', 'internship', 'role', 'candidate',
    'interview', 'offer', 'contract'  # Add more keywords here
]
```

## Troubleshooting

### "Error: 'credentials.json' not found"
- Make sure you downloaded the OAuth credentials from Google Cloud Console
- File must be named exactly `credentials.json`
- It should be in the same folder as `job_tracker.py`

### "Error: The caller does not have permission to access the requested resource"
- You may not have enabled the Gmail API
- Go back to [Google Cloud Console](https://console.cloud.google.com/)
- Search "Gmail API" and click "ENABLE"

### "ModuleNotFoundError: No module named 'google...'"
- Install dependencies: `pip install -r requirements.txt`

### Excel file won't open
- Make sure no other program has the file open
- Try deleting `job_applications.xlsx` and running the script again

### No emails found
- Check your Gmail inbox has recruitment emails
- Try running with different search query: `query='is:all'` to scan everything
- Verify keywords in `JOB_KEYWORDS` match your emails

## Project Structure

```
job-tracker-automator/
├── job_tracker.py              # Main application (connects to Gmail API)
├── requirements.txt            # Python dependencies
├── README.md                   # This file
├── .gitignore                  # Prevents credentials from being uploaded
├── credentials.json            # OAuth credentials (NEVER commit this!)
├── token.json                  # Auth token (auto-generated, don't commit)
└── job_applications.xlsx       # Output spreadsheet (auto-generated)
```

## How It Works

1. **Authenticates** with your Gmail using OAuth 2.0 (secure, no password needed)
2. **Searches** your inbox for job-related emails using keywords
3. **Extracts** sender, subject, date, and email body
4. **Filters** emails that match job keywords
5. **Organizes** data by company, position, and date
6. **Exports** to Excel with professional formatting and color coding

## Security & Privacy

- **Your credentials are safe**: Uses OAuth 2.0 (industry standard)
- **No passwords stored**: Only authorization token is saved locally
- **Read-only access**: App can only read emails, never modify them
- **Local processing**: All data stays on your computer
- **Git protection**: `.gitignore` ensures credentials never upload to GitHub

## Contributing

Found a bug or want to improve this? 

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Commit (`git commit -m 'Add amazing feature'`)
5. Push to your branch (`git push origin feature/amazing-feature`)
6. Open a Pull Request

## License

This project is open source and available under the MIT License.

## Author

**Anthony Jaldin Guzman**  
George Mason University - Cybersecurity Engineering Student  
📧 a.jaldin18@gmail.com  
🔗 [LinkedIn](https://linkedin.com/in/anthonyguzman)

---

## Changelog

### v1.0.0 (Initial Release)
- Gmail API integration
- Email keyword filtering
- Excel export with formatting
- Color-coded status tracking

## FAQ

**Q: Will this affect my Gmail account?**  
A: No, it only reads emails. Your account remains unchanged.

**Q: How often should I run this?**  
A: Run it weekly to stay organized. You can even set up a scheduled task.

**Q: Can I modify the Excel output after creation?**  
A: Yes! The script generates a fresh file each time, so manually edit the file and re-run when you have new emails.

**Q: What if I have thousands of emails?**  
A: The script processes up to 20 emails by default. Increase `max_results` in the code if needed.

---

**Need help?** Check the Troubleshooting section above or open a GitHub Issue!
