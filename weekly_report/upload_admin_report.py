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


def upload_admin_report(*, drive_service: Any, admin_report_to_upload: str | Path, folder_id: str,
                        email_list: list[str], send_email_flag: bool, weekly_msg: str | None,
                        monthly_msg: str | None) -> None:
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
    # delete_file_list1 = get_google_file_ids(drive_service, file_name_wo_ext, folder_id)
    delete_file_list = get_google_file_or_folder_ids(drive_service, 'file', file_name_wo_ext, folder_id)
    delete_list_of_google_files(drive_service, delete_file_list)

    print('Uploading filename ', file_name_wo_ext)
    uploaded_file_id = upload_sheet_to_drive(drive_service, admin_report_to_upload, folder_id)

    # set permissions and email admin report to Core group
    if not send_email_flag:
        permission_msg = None
    elif file_name_wo_ext[-2:].lower() == '-m':
        permission_msg = monthly_msg
        # permission_msg = (f"A new MONTHLY ROV-WIDE Sincere summary report has been sent to the CORE GROUP. "
        #                   f"Click to open.")
    else:
        permission_msg = weekly_msg
        # permission_msg = (f"A new WEEKLY ROV-WIDE Sincere summary report has been sent to the CORE GROUP. "
        #                   f"Click to open.")

    for email in email_list:
        # print('  - Adding ROV-wide permission for email: ', email)
        # logging.getLogger(name='my_logger').info(f"{inspect.stack()[0][3]}  - Adding ROV-wide permission for email: "
        #                                          f"', {email}")
        logger.debug(f"Adding ROV-wide permission for email: {email}")
        permission_to_drive_file(drive_service, uploaded_file_id, send_email_flag, email, permission_msg)
