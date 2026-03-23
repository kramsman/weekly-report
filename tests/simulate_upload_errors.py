"""Simulate a permission failure and a SendGrid failure to verify error logging + popup.

Run with:
    .venv/bin/python tests/simulate_upload_errors.py
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pathlib import Path
from unittest.mock import MagicMock, patch

from uvbekutils import pyautobek
from weekly_report.constants import ERROR_LOG_FILE
from weekly_report.send_notification_email import send_drive_notification
from weekly_report.upload_admin_report import upload_admin_report

# ── setup ──────────────────────────────────────────────────────────────────
# clear any leftover error log
if ERROR_LOG_FILE.exists():
    ERROR_LOG_FILE.unlink()
    print(f"Cleared old error log: {ERROR_LOG_FILE}")

# fake api key file
api_key_file = Path("/tmp/fake_sendgrid_key.txt")
api_key_file.write_text("bad-key-intentional")

fake_report = Path("/tmp/ROVWide Summary 2026-03-23-W.xlsx")
fake_report.write_bytes(b"fake")

# ── simulate permission failure ────────────────────────────────────────────
print("\n[1] Simulating permission failure...")

with (
    patch("weekly_report.upload_admin_report.get_google_file_or_folder_ids", return_value=[]),
    patch("weekly_report.upload_admin_report.delete_list_of_google_files"),
    patch("weekly_report.upload_admin_report.upload_sheet_to_drive", return_value="fake-file-id"),
    patch("weekly_report.upload_admin_report.permission_to_drive_file",
          side_effect=Exception("403 The user does not have sufficient permissions")),
    patch("weekly_report.upload_admin_report.send_drive_notification"),  # skip email for this step
):
    upload_admin_report(
        drive_service=MagicMock(),
        admin_report_to_upload=fake_report,
        folder_id="fake-folder-id",
        email_list=["blocked@example.com"],
        send_email_flag=False,
        weekly_msg="Test msg",
        weekly_subject="Test subject",
        monthly_msg=None,
        monthly_subject=None,
        sendgrid_api_key_file=api_key_file,
        sendgrid_from_email="sender@example.com",
    )

print(f"   Error log exists: {ERROR_LOG_FILE.exists()}")
if ERROR_LOG_FILE.exists():
    print(f"   Contents so far:\n{ERROR_LOG_FILE.read_text()}")

# ── simulate bad SendGrid key ──────────────────────────────────────────────
print("\n[2] Simulating SendGrid failure (bad API key)...")

send_drive_notification(
    to_email="org@example.com",
    file_id="fake-file-id",
    message="Hello organizer",
    subject="Weekly Report — Test Room",
    from_email="sender@example.com",
    api_key_file=api_key_file,   # bad key — will raise
    error_log_file=ERROR_LOG_FILE,
)

print(f"   Error log exists: {ERROR_LOG_FILE.exists()}")
if ERROR_LOG_FILE.exists():
    print(f"   Final contents:\n{ERROR_LOG_FILE.read_text()}")

# ── show the summary popup ─────────────────────────────────────────────────
print("\n[3] Showing alert popup...")
if ERROR_LOG_FILE.exists():
    pyautobek.alert_with_file_link(
        "Errors occurred during upload. See details:",
        ERROR_LOG_FILE,
        "Upload Errors"
    )
else:
    print("   No error log found — nothing to show.")
