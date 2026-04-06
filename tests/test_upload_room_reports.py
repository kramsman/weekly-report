"""Tests for error-log-file behavior in upload_room_reports."""

from contextlib import ExitStack
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from weekly_report.upload_room_reports import upload_room_reports


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _make_room_dir(tmp_path: Path, room_name: str, monthly: bool = False) -> Path:
    """Create a temp upload dir containing one fake xlsx file for room_name."""
    upload_dir = tmp_path / "VL Org Reports"
    upload_dir.mkdir()
    suffix = "-M" if monthly else "-W"
    filename = f"{room_name} Sincere Request Summary 2026-03-23{suffix}.xlsx"
    (upload_dir / filename).write_bytes(b"fake xlsx")
    return upload_dir


def _enter_base_patches(stack: ExitStack, error_log_file: Path) -> None:
    """Enter the common patches needed to run upload_room_reports without hitting Google APIs."""
    stack.enter_context(patch("weekly_report.upload_room_reports.get_google_file_or_folder_ids", return_value=[]))
    stack.enter_context(patch("weekly_report.upload_room_reports.delete_list_of_google_files"))
    stack.enter_context(patch("weekly_report.upload_room_reports.create_drive_subfolder", return_value="new-folder-id"))
    stack.enter_context(patch("weekly_report.upload_room_reports.upload_sheet_to_drive", return_value="uploaded-file-id"))
    stack.enter_context(patch("weekly_report.upload_room_reports.send_drive_notification"))
    stack.enter_context(patch("weekly_report.upload_room_reports.ERROR_LOG_FILE", error_log_file))


# ─────────────────────────────────────────────────────────────────────────────
# Production path (organizer_email_list matching)
# ─────────────────────────────────────────────────────────────────────────────

def test_upload_room_reports_logs_permission_failure(tmp_path, api_key_file, error_log_file):
    """When permission_to_drive_file raises, PERMISSION ERROR is written to error_log_file."""
    room_name = "Springfield Volunteers"
    org_name = "Springfield Volunteers"  # no comma, so replace(",","-") is a no-op
    email = "org@example.com"
    upload_dir = _make_room_dir(tmp_path, room_name)

    print(f"\n--- test_upload_room_reports_logs_permission_failure ---")
    print(f"  Parameters : room={room_name!r}, email={email!r}")
    print(f"  Input      : permission_to_drive_file raises RuntimeError")

    with ExitStack() as stack:
        _enter_base_patches(stack, error_log_file)
        stack.enter_context(patch("weekly_report.upload_room_reports.permission_to_drive_file",
                                  side_effect=RuntimeError("403 Insufficient permission")))

        upload_room_reports(
            drive_service=MagicMock(),
            str_dir_to_upload=str(upload_dir),
            organizer_email_list=[[org_name, email]],
            folder_id="parent-folder-id",
            send_email_flag=False,
            weekly_msg="Weekly msg",
            weekly_subject="Weekly subject",
            monthly_msg=None,
            monthly_subject=None,
            sendgrid_api_key_file=api_key_file,
            sendgrid_from_email="sender@example.com",
        )

    assert error_log_file.exists(), "error_log_file should be created on permission failure"
    contents = error_log_file.read_text()
    print(f"  Output     : error_log contents={contents!r}")
    assert f"PERMISSION ERROR: room: {room_name}" in contents
    assert f"email: {email}" in contents
    assert "403 Insufficient permission" in contents


def test_upload_room_reports_no_log_on_success(tmp_path, api_key_file, error_log_file):
    """When permission succeeds, error_log_file is not created."""
    room_name = "Springfield Volunteers"
    upload_dir = _make_room_dir(tmp_path, room_name)

    print(f"\n--- test_upload_room_reports_no_log_on_success ---")
    print(f"  Parameters : room={room_name!r}")
    print(f"  Input      : permission_to_drive_file succeeds")

    with ExitStack() as stack:
        _enter_base_patches(stack, error_log_file)
        stack.enter_context(patch("weekly_report.upload_room_reports.permission_to_drive_file"))

        upload_room_reports(
            drive_service=MagicMock(),
            str_dir_to_upload=str(upload_dir),
            organizer_email_list=[["Springfield Volunteers", "org@example.com"]],
            folder_id="parent-folder-id",
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


def test_upload_room_reports_logs_comma_room_name(tmp_path, api_key_file, error_log_file):
    """Organizer names with commas are matched after replacing ',' with '-' in the filename."""
    # Org name "Smith, Jane" → file is named "Smith- Jane Sincere ..."
    org_name = "Smith, Jane"
    room_name = "Smith- Jane"  # comma replaced with dash in filename
    email = "jane@example.com"
    upload_dir = _make_room_dir(tmp_path, room_name)

    print(f"\n--- test_upload_room_reports_logs_comma_room_name ---")
    print(f"  Parameters : org={org_name!r}, room_name_in_file={room_name!r}, email={email!r}")
    print(f"  Input      : permission_to_drive_file raises")

    with ExitStack() as stack:
        _enter_base_patches(stack, error_log_file)
        stack.enter_context(patch("weekly_report.upload_room_reports.permission_to_drive_file",
                                  side_effect=RuntimeError("blocked")))

        upload_room_reports(
            drive_service=MagicMock(),
            str_dir_to_upload=str(upload_dir),
            organizer_email_list=[[org_name, email]],
            folder_id="parent-folder-id",
            send_email_flag=False,
            weekly_msg="msg",
            weekly_subject="subj",
            monthly_msg=None,
            monthly_subject=None,
            sendgrid_api_key_file=api_key_file,
            sendgrid_from_email="sender@example.com",
        )

    assert error_log_file.exists()
    contents = error_log_file.read_text()
    print(f"  Output     : error_log contents={contents!r}")
    assert f"room: {room_name}" in contents
    assert f"email: {email}" in contents


# ─────────────────────────────────────────────────────────────────────────────
# Test mode (test_email_list path)
# ─────────────────────────────────────────────────────────────────────────────

def test_upload_room_reports_test_mode_logs_permission_failure(tmp_path, api_key_file, error_log_file):
    """In test mode (test_email_list), permission failure is still logged to error_log_file."""
    room_name = "Springfield Volunteers"
    test_emails = ["tester@example.com"]
    upload_dir = _make_room_dir(tmp_path, room_name)

    print(f"\n--- test_upload_room_reports_test_mode_logs_permission_failure ---")
    print(f"  Parameters : room={room_name!r}, test_emails={test_emails}")
    print(f"  Input      : permission_to_drive_file raises in test mode")

    with ExitStack() as stack:
        _enter_base_patches(stack, error_log_file)
        stack.enter_context(patch("weekly_report.upload_room_reports.permission_to_drive_file",
                                  side_effect=RuntimeError("403 Insufficient permission")))

        upload_room_reports(
            drive_service=MagicMock(),
            str_dir_to_upload=str(upload_dir),
            organizer_email_list=[],          # ignored in test mode
            folder_id="parent-folder-id",
            send_email_flag=False,
            weekly_msg="Weekly msg",
            weekly_subject="Weekly subject",
            monthly_msg=None,
            monthly_subject=None,
            sendgrid_api_key_file=api_key_file,
            sendgrid_from_email="sender@example.com",
            test_email_list=test_emails,
        )

    assert error_log_file.exists(), "error_log_file should be created on permission failure in test mode"
    contents = error_log_file.read_text()
    print(f"  Output     : error_log contents={contents!r}")
    assert f"email: {test_emails[0]}" in contents
    assert "403 Insufficient permission" in contents
