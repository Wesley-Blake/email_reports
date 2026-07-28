# email_reports

This script collects CSV attachments from an Outlook folder named email_reports, saves them to your Downloads/email_reports folder, and reviews the files to identify reports that need attention.

## What it does
- Clears the local report folder before each run.
- Looks through Outlook messages in the email_reports subfolder of your Inbox.
- Saves any attached CSV files locally.
- Reviews each CSV file for a column whose name contains camp and keeps the files that contain a value of 2.
- Removes files that do not meet the criteria and prints a summary of the remaining reports.

## Requirements
- Python 3
- pandas
- pywin32 (Windows only)
- Microsoft Outlook installed and available on the system

## Usage
1. Create an Outlook subfolder named email_reports under your Inbox.
2. Place email messages containing CSV attachments into that folder.
3. Run the script:

   python email_reports.py

> The script uses Outlook Automation via COM, so it is intended for Windows environments with Outlook installed.
