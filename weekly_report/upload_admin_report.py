"""Upload the admin (ROV-wide) report to Google Drive and notify the Core group."""

import inspect
import logging
from pathlib import Path
from typing import Any
from loguru import logger

from bekgoogle.delete_list_of_google_files import delete_list_of_google_files
from bekgoogle.get_google_file_or_folder_ids import get_google_file_or_folder_ids
from bekgoogle.permission_to_drive_file import permission_to_drive_file
from bekgoogle.upload_sheet_to_drive import upload_sheet_to_drive
from weekly_report.constants import ERROR_LOG_FILE
from weekly_report.send_notification_email import send_drive_notification


def upload_admin_report(*, drive_service: Any, admin_report_to_upload: str | Path, folder_id: str,
                        email_list: list[str], send_email_flag: bool,
                        weekly_msg: str | None, weekly_subject: str | None,
                        monthly_msg: str | None, monthly_subject: str | None,
                        sendgrid_api_key_file: Path, sendgrid_from_email: str) -> None:
    """Upload the admin (ROV-wide) report to Google Drive and notify the Core group.

    Deletes any existing Drive file with the same name in the admin folder
    before uploading, then grants reader permission to every address in
    email_list.

    Args:
        * ():
        drive_service: Authenticated Google Drive API service resource.
        admin_report_to_upload: Full path to the local admin .xlsx file.
        folder_id: Google Drive folder ID for the admin report destination.
        email_list: Email addresses to grant reader permission.
        send_email_flag: Whether to send Google notification emails.
        weekly_msg: Notification message body for weekly reports.
        monthly_msg: Notification message body for monthly reports.
    """
    # do the upload
    admin_report_to_upload = Path(admin_report_to_upload)
    file_name_wo_ext = Path(admin_report_to_upload).stem  # filename without extension

    # delete all files with same name in the destination admin folder
    delete_file_list = get_google_file_or_folder_ids(drive_service, 'file', file_name_wo_ext, folder_id)
    delete_list_of_google_files(drive_service, delete_file_list)

    print('Uploading filename ', file_name_wo_ext)
    uploaded_file_id = upload_sheet_to_drive(drive_service, admin_report_to_upload, folder_id)

    # determine notification message based on report type
    if file_name_wo_ext[-2:].lower() == '-m':
        notification_msg = monthly_msg
        notification_subject = monthly_subject
    else:
        notification_msg = weekly_msg
        notification_subject = weekly_subject

    for email in email_list:
        logger.debug(f"Adding ROV-wide permission for email: {email}")
        # suppress Google's notification — SendGrid sends a custom email instead
        try:
            permission_to_drive_file(drive_service, uploaded_file_id, False, email, None)
        except Exception as e:
            logger.error(f"Permission failed: email: {email}: {e}")
            with open(ERROR_LOG_FILE, 'a') as f:
                f.write(f"PERMISSION ERROR: email: {email}, error: {e}\n")
        if send_email_flag and notification_msg:
            send_drive_notification(email, uploaded_file_id, notification_msg, notification_subject,
                                    sendgrid_from_email, sendgrid_api_key_file,
                                    all_recipients=email_list,
                                    error_log_file=ERROR_LOG_FILE)
