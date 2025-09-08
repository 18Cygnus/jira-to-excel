# Jira to Excel Exporter

Export Jira issues from a saved filter to a formatted Excel workbook:
- Date/time columns formatted like Jira (created/updated, due_date)
- Excel **Table** with **filter dropdowns** and sortable date columns
- Optional: save outputs to `./excels/`

## Requirements
- Python 3.9+ (recommended)
- A Jira Cloud user with access to the target filter
- API Token from Atlassian (https://id.atlassian.com/manage-profile/security/api-tokens)

## Setup

### 1. Clone Repo
```bash
git clone https://github.com/18Cygnus/jira-to-excel.git
cd jira-to-excel
```
### 2. Create .venv
Create virtual environment in local
```bash
python -m venv .venv
```
### 3. Activate environment
- Windows:
```bash
source .venv/Scripts/activate
```
- Mac/Linux:
```bash
source .venv/bin/activate
```
### 4. Install dependencies
```bash
pip install -r requirements.txt
```
### 5. Configure .env
There is `.env.example` within the repository you can look into
```bash
JIRA_BASE_URL=https://yourcompany.atlassian.net

# Fill in your Jira email that have access/authority to your desired filter
JIRA_EMAIL=you@company.com

# Configure your account's API token in https://id.atlassian.com/manage-profile/security/api-tokens
JIRA_API_TOKEN=your_api_token

# Use one of these:
# Filter ID appears at the end of the URL path "https://company.atlassian.net/issues/?filter=xxxxx"
JIRA_FILTER_ID=12345
# or
JIRA_FILTER_NAME=FilterName

```
### 6. Run the script
```bash
python jira_export.py
```

## Scheduler Setup (schtasks) - Windows

> [!WARNING]  
> Please refrain from opening the exported Excel (.xlsx) file when the task is currently running to avoid export failure

### 1. Open Command Line (CMD) in Administrator Mode
Make sure to adjust the path with your own local path.
```bash
schtasks /Create ^
  /SC DAILY /ST 11:00 ^
  /TN "Jira Export to Excel" ^
  /TR "cmd /c \"cd /d C:\yourdirectory\jira-to-excel && C:\yourdirectory\jira-to-excel\.venv\Scripts\python.exe jira_export.py >> excels\run.log 2>&1\"" ^
  /RL HIGHEST
```
- `/SC DAILY /ST 11:00` This will schedule the export at 11:00 daily.
- **Option 1** is more reliable and easier to debug
- **Option 2** uses PowerShell for more complex command handling

### 2. Running or deleting the task
Running the task :
```bash
schtasks /Run /TN "Jira Export to Excel"
```

Deleting the task if necessary:
```bash
schtasks /Delete /TN "Jira Export to Excel" /F
```

### 3. Handling Missed Schedules (GUI)
Windows Task Scheduler will not run a task if the computer is off at the scheduled time. For example, if your task is set at 11:00, but your computer is powered off at that time and turned back on at 12:00, the task will be skipped and will only run again at the next scheduled time (the next day at 11:00).

### Run Task After a Missed Schedule :
1. Search and open Task Scheduler (`taskschd.msc`).
2. Find your task (e.g., "*Jira Export to Excel*").
3. Right click → **Properties** → **Settings tab**.
4. Enable “Run task as soon as possible after a scheduled start is missed” option.

With this setting enabled, if the system is off at the scheduled time, the task will automatically run the next time the computer is turned on.

# Google Spreadsheet Integration Setup  
### 1. Activate Google Sheets API
1. Open [Google Cloud Console](https://console.cloud.google.com/)
2. Select project or create new and configure: <br/>- Fill in "Project Name" <br/>- Leave "Location/Organization" **empty**
3. Go to **APIs & Services** → **Library**
4. Find **Google Sheets API** → **Enable**

### 2. Create Service Account
1. In **APIs & Services** → click **Create Credentials** → choose **Application Data** → click Next
2. Fill in **Name** (e.g: jira-export-service)
3. After service account is created, open **Keys** tab → **Add Key** → **Create new key**.
4. Choose **JSON** & download file <br/>- Save this file in local, e.g: `C:\directory\jira-to-excel\service_account.json` <br/>- Make sure to include this file inside **.gitignore**

### 3. Create Spreadsheet & Grant Access to Google Sheet
> [!IMPORTANT]  
> Please start from a fresh/new spreadsheet
1. Go to [Spreadsheet](https://docs.google.com/spreadsheets/u/0/) and create **New Spreadsheet**
2. Open Spreadsheet e.g:
`https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit`
3. Click Share & Add Service Account email e.g: `jira-export-service@my-project.iam.gserviceaccount.com` as **Editor**. Find the email in [Service Account Tab](https://console.cloud.google.com/iam-admin/serviceaccounts)

### 4. Update `.ENV`
Add these variables in `.env`:
```bash
GSHEET_ID=<SPREADSHEET_ID>
GSHEET_WORKSHEET=Issues
GOOGLE_SERVICE_ACCOUNT_FILE=C:\YourDirectory\jira-to-excel\service_account.json
```
- GSHEET_ID = ID from spreadsheet's URL (between `/d/` and `/edit`)
- GSHEET_WORKSHEET = worksheet tab's name (e.g: Issues)
- GOOGLE_SERVICE_ACCOUNT_FILE = path to JSON key service account file

### 5. Install Dependencies
Inside your virtual environment (`.venv`), run:
```bash
pip install -r requirements.txt
```

### 6. Run the Exporter
Run the script manually or wait for the scheduler. <br/>
This python script will:
- Fetch issues from Jira
- Export to excels/jira_export.xlsx (`local`)
- Overwrite Google Sheets' worksheet (`cloud`)