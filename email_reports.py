"""Collect, save, and review CSV report attachments from Outlook."""

from datetime import datetime
from pathlib import Path

import pandas as pd

try:
    import win32com.client
except ImportError:  # pragma: no cover - depends on the runtime environment
    WIN32COM = None

SAVE_DIR = Path.home() / "Downloads" / "email_reports"
OUTLOOK_FOLDER_NAME = "email_reports"
TODAY = datetime.now().date()


def clear_reports_dir() -> None:
    """Create the report directory and remove any existing files from it."""
    SAVE_DIR.mkdir(parents=True, exist_ok=True)
    for file in SAVE_DIR.iterdir():
        if file.is_file():
            file.unlink()


def _get_outlook_subfolder():
    """Return the Outlook subfolder used to read report emails."""
    if win32com is None:
        raise RuntimeError("pywin32 is required to access Outlook on Windows.")

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)

    try:
        return inbox.Folders[OUTLOOK_FOLDER_NAME]
    except Exception as exc:  # pragma: no cover - depends on Outlook configuration
        raise RuntimeError(
            f"Could not find Outlook folder '{OUTLOOK_FOLDER_NAME}'."
        ) from exc


def extract_attention_rows(df: pd.DataFrame) -> pd.DataFrame | None:
    """Return rows that should be kept for reports requiring attention."""
    for header in df.columns:
        if "camp" not in header.lower():
            continue

        attention_rows = df[df[header] == 2]
        if not attention_rows.empty:
            return attention_rows

    return None


def collect_csv() -> None:
    """Read Outlook messages from the report folder and save CSV attachments."""
    subfolder = _get_outlook_subfolder()

    for message in list(subfolder.Items):
        if message.ReceivedTime.date() < TODAY:
            message.Delete()
            continue

        attachments = message.Attachments
        if attachments.Count == 0:
            message.Delete()
            continue

        saved_any_csv = False
        for index in range(1, attachments.Count + 1):
            attachment = attachments.Item(index)
            if attachment.FileName.lower().endswith(".csv"):
                attachment.SaveAsFile(SAVE_DIR / attachment.FileName)
                saved_any_csv = True

        if not saved_any_csv:
            message.Delete()


def check_csv() -> None:
    """Review each saved CSV file and keep only reports that require attention."""
    total = 0
    for file in SAVE_DIR.glob("*.csv"):
        try:
            df = pd.read_csv(file)
        except pd.errors.EmptyDataError:
            file.unlink()
            continue

        attention_rows = extract_attention_rows(df)
        if attention_rows is None:
            file.unlink()
            continue

        attention_rows.to_csv(file, index=False)
        total += 1

    print(f"Total number of reports need attention: {total}")


def main():
    """Run the report collection and validation workflow."""
    clear_reports_dir()
    collect_csv()
    check_csv()


if __name__ == "__main__":
    main()
