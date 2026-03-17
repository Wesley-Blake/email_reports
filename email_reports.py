from pathlib import Path
from datetime import datetime
import win32com.client
import pandas as pd


SAVE_DIR = Path.home() / 'Downloads' / 'email_reports'

def clear_reports_dir() -> None:
    SAVE_DIR.mkdir(parents=True, exist_ok=True)
    for file in SAVE_DIR.iterdir():
        file.unlink()

def collect_csv() -> None:
    TODAY = datetime.now().date()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    inbox = namespace.GetDefaultFolder(6)
    subfolder = inbox.Folders['email_reports']

    index = 0
    messages = subfolder.Items
    while index < len(messages):
        message = messages[index]
        # These messages are just on repeat, I don't want the same document over and again.
        if message.ReceivedTime.date() < TODAY:
            message.Delete()
            messages = subfolder.Items
            continue

        attachment_count = message.Attachments.Count
        if attachment_count > 0:
            for i in range(1, attachment_count + 1):
                attachment = message.Attachments.Item(i)
                if attachment.FileName.endswith('.csv'):
                    file_path = SAVE_DIR / attachment.FileName
                    attachment.SaveAsFile(file_path)
        else:
            message.Delete()
            messages = subfolder.Items
            continue
        index += 1

def check_csv() -> None:
    total = 0
    for file in SAVE_DIR.iterdir():
        delete = False
        df = pd.read_csv(file)
        headers = df.columns
        for header in headers:
            if 'camp' in header.lower():
                if any(df[header] == 2):
                    df[df[header] == 2].to_csv(file, index=False)
                    break
                else:
                    delete = True
                    break
        if delete:
            file.unlink()
        else:
            total += 1
    print(f"Total number of reports need attention: {total}")

def main():
    clear_reports_dir()
    collect_csv()
    check_csv()


if __name__ == '__main__':
    main()
