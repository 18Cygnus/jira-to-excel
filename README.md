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
### 4. Install dependency
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

## Scheduler Setup (schtasks)

### 1. Open Command Line (CMD) in Administrator Mode
Make sure to adjust the path with your own local path.
```bash
schtasks /Create ^
  /SC DAILY /ST 11:00 ^
  /TN "Jira Export to Excel" ^
  /TR "\"C:\yourdirectory\jira-to-excel\.venv\Scripts\python.exe\" \"C:\yourdirectory\jira-to-excel\jira_export.py\" >> \"C:\yourdirectory\jira-to-excel\excels\run.log\" 2>&1" ^
  /RL HIGHEST
```
- `/SC MINUTE /ST 11 ^` This will schedule the export at 11:00 WIB daily.

### 2. Running or deleting the task
Running the task :
```bash
schtasks /Run /TN "Jira Export to Excel"
```

Deleting the task if necessary:
### 
```bash
schtasks /Delete /TN "Jira Export to Excel" /F
```