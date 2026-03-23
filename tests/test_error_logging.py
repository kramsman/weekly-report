# SKIPPED (too simple to test):
#   - send_drive_notification html construction: pure string formatting, no logic

"""Tests for error-log-file behavior in send_notification_email and upload_admin_report."""

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from weekly_report.send_notification_email import send_drive_notification


# ─────────────────────────────────────────────────────────────────────────────
# send_drive_notification
# ─────────────────────────────────────────────────────────────────────────────

def test_send_drive_notification_logs_on_sendgrid_failure(api_key_file, error_log_file):
    """When SendGrid raises, the error is appended to error_log_file."""
    # Arrange
    to_email = "org@example.com"
    subject = "Weekly Report"
    file_id = "abc123"

    print(f"\n--- test_send_drive_notification_logs_on_sendgrid_failure ---")
    print(f"  Parameters : to_email={to_email!r}, subject={subject!r}, file_id={file_id!r}")
    print(f"  Input      : api_key_file={api_key_file}, error_log_file={error_log_file}")

    with patch("weekly_report.send_notification_email.sendgrid.SendGridAPIClient") as mock_sg_cls:
        mock_sg_cls.return_value.send.side_effect = Exception("401 Unauthorized")

        # Act
        send_drive_notification(
            to_email=to_email,
            file_id=file_id,
            message="Hello organizer",
            subject=subject,
            from_email="sender@example.com",
            api_key_file=api_key_file,
            error_log_file=error_log_file,
        )

    # Assert
    assert error_log_file.exists(), "error_log_file should be created on failure"
    contents = error_log_file.read_text()
    print(f"  Output     : error_log contents={contents!r}")
    assert f"EMAIL ERROR: email: {to_email}" in contents
    assert f"subject: {subject}" in contents
    assert "401 Unauthorized" in contents


def test_send_drive_notification_no_log_on_success(api_key_file, error_log_file):
    """When SendGrid succeeds, error_log_file is not created."""
    # Arrange
    to_email = "org@example.com"
    subject = "Weekly Report"

    print(f"\n--- test_send_drive_notification_no_log_on_success ---")
    print(f"  Parameters : to_email={to_email!r}, subject={subject!r}")
    print(f"  Input      : api_key_file={api_key_file}, error_log_file={error_log_file}")

    mock_response = MagicMock()
    mock_response.status_code = 202

    with patch("weekly_report.send_notification_email.sendgrid.SendGridAPIClient") as mock_sg_cls:
        mock_sg_cls.return_value.send.return_value = mock_response

        # Act
        send_drive_notification(
            to_email=to_email,
            file_id="abc123",
            message="Hello organizer",
            subject=subject,
            from_email="sender@example.com",
            api_key_file=api_key_file,
            error_log_file=error_log_file,
        )

    # Assert
    print(f"  Output     : error_log_file exists={error_log_file.exists()}")
    assert not error_log_file.exists(), "error_log_file should NOT be created on success"


def test_send_drive_notification_no_log_file_param(api_key_file):
    """When error_log_file is None and SendGrid fails, no file is written (no crash)."""
    # Arrange
    print(f"\n--- test_send_drive_notification_no_log_file_param ---")
    print(f"  Parameters : error_log_file=None")
    print(f"  Input      : SendGrid raises, no error_log_file passed")

    with patch("weekly_report.send_notification_email.sendgrid.SendGridAPIClient") as mock_sg_cls:
        mock_sg_cls.return_value.send.side_effect = Exception("timeout")

        # Act — should not raise
        send_drive_notification(
            to_email="org@example.com",
            file_id="abc123",
            message="Hello",
            subject="Test",
            from_email="sender@example.com",
            api_key_file=api_key_file,
            error_log_file=None,
        )

    print(f"  Output     : no exception raised, no file written")


# ─────────────────────────────────────────────────────────────────────────────
# upload_admin_report — permission failure logging
# ─────────────────────────────────────────────────────────────────────────────

def test_upload_admin_report_logs_permission_failure(tmp_path, api_key_file, error_log_file):
    """When permission_to_drive_file raises, PERMISSION ERROR is written to ERROR_LOG_FILE."""
    from weekly_report.upload_admin_report import upload_admin_report

    # Arrange — create a fake xlsx file (only stem name matters; upload is mocked)
    fake_report = tmp_path / "ROVWide Summary 2026-03-23-W.xlsx"
    fake_report.write_bytes(b"fake xlsx content")

    email = "core@example.com"
    print(f"\n--- test_upload_admin_report_logs_permission_failure ---")
    print(f"  Parameters : admin_report={fake_report.name}, email={email!r}")
    print(f"  Input      : permission_to_drive_file raises RuntimeError")

    drive_service = MagicMock()

    with (
        patch("weekly_report.upload_admin_report.get_google_file_or_folder_ids", return_value=[]),
        patch("weekly_report.upload_admin_report.delete_list_of_google_files"),
        patch("weekly_report.upload_admin_report.upload_sheet_to_drive", return_value="file-id-xyz"),
        patch("weekly_report.upload_admin_report.permission_to_drive_file",
              side_effect=RuntimeError("blocked by recipient")),
        patch("weekly_report.upload_admin_report.send_drive_notification"),
        patch("weekly_report.upload_admin_report.ERROR_LOG_FILE", error_log_file),
    ):
        # Act
        upload_admin_report(
            drive_service=drive_service,
            admin_report_to_upload=fake_report,
            folder_id="folder-abc",
            email_list=[email],
            send_email_flag=False,
            weekly_msg="Weekly msg",
            weekly_subject="Weekly subject",
            monthly_msg=None,
            monthly_subject=None,
            sendgrid_api_key_file=api_key_file,
            sendgrid_from_email="sender@example.com",
        )

    # Assert
    assert error_log_file.exists(), "error_log_file should be created on permission failure"
    contents = error_log_file.read_text()
    print(f"  Output     : error_log contents={contents!r}")
    assert f"PERMISSION ERROR: email: {email}" in contents
    assert "blocked by recipient" in contents


def test_upload_admin_report_no_log_on_success(tmp_path, api_key_file, error_log_file):
    """When permission succeeds, ERROR_LOG_FILE is not created."""
    from weekly_report.upload_admin_report import upload_admin_report

    # Arrange
    fake_report = tmp_path / "ROVWide Summary 2026-03-23-W.xlsx"
    fake_report.write_bytes(b"fake xlsx content")

    print(f"\n--- test_upload_admin_report_no_log_on_success ---")
    print(f"  Parameters : admin_report={fake_report.name}")
    print(f"  Input      : permission_to_drive_file succeeds")

    drive_service = MagicMock()

    with (
        patch("weekly_report.upload_admin_report.get_google_file_or_folder_ids", return_value=[]),
        patch("weekly_report.upload_admin_report.delete_list_of_google_files"),
        patch("weekly_report.upload_admin_report.upload_sheet_to_drive", return_value="file-id-xyz"),
        patch("weekly_report.upload_admin_report.permission_to_drive_file"),
        patch("weekly_report.upload_admin_report.send_drive_notification"),
        patch("weekly_report.upload_admin_report.ERROR_LOG_FILE", error_log_file),
    ):
        # Act
        upload_admin_report(
            drive_service=drive_service,
            admin_report_to_upload=fake_report,
            folder_id="folder-abc",
            email_list=["core@example.com"],
            send_email_flag=False,
            weekly_msg="Weekly msg",
            weekly_subject="Weekly subject",
            monthly_msg=None,
            monthly_subject=None,
            sendgrid_api_key_file=api_key_file,
            sendgrid_from_email="sender@example.com",
        )

    print(f"  Output     : error_log_file exists={error_log_file.exists()}")
    assert not error_log_file.exists(), "error_log_file should NOT be created on success"
