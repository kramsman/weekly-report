"""Upload the admin (ROV-wide) report to Google Drive and notify the Core group."""

import inspect
import logging
from pathlib import Path
from typing import Any

from weekly_report.constants import ADMIN_REPORT_FOLDER_ID
from weekly_report.constants import CORE_EMAIL_LIST
from weekly_report.constants import CORE_MONTHLY_MSG
from weekly_report.constants import CORE_WEEKLY_MSG
from weekly_report.constants import SEND_PERMISSION_EMAIL_FLAG
from weekly_report.google_scripts.delete_list_of_google_files import delete_list_of_google_files
from weekly_report.google_scripts.get_google_file_or_folder_ids import get_google_file_or_folder_ids
from weekly_report.google_scripts.permission_to_drive_file import permission_to_drive_file
from weekly_report.google_scripts.upload_sheet_to_drive import upload_sheet_to_drive


def upload_admin_report(drive_service: Any, admin_report_to_upload: str | Path) -> None:
    """Upload the admin (ROV-wide) report to Google Drive and notify the Core group.

    Deletes any existing Drive file with the same name in the admin folder
    before uploading, then grants reader permission to every address in
    CORE_EMAIL_LIST.

    Args:
        drive_service: Authenticated Google Drive API service resource.
        admin_report_to_upload: Full path to the local admin .xlsx file.
    """
    # do the upload
    admin_report_to_upload = Path(admin_report_to_upload)
    file_name_wo_ext = Path(admin_report_to_upload).stem  # filename without extension

    # delete all files with same name in the destination admin folder
    # delete_file_list1 = get_google_file_ids(drive_service, file_name_wo_ext, ADMIN_REPORT_FOLDER_ID)
    delete_file_list = get_google_file_or_folder_ids(drive_service, 'file', file_name_wo_ext, ADMIN_REPORT_FOLDER_ID)
    delete_list_of_google_files(drive_service, delete_file_list)

    print('Uploading filename ', file_name_wo_ext)
    uploaded_file_id = upload_sheet_to_drive(drive_service, admin_report_to_upload, ADMIN_REPORT_FOLDER_ID)

    # set permissions and email admin report to Core group
    if not SEND_PERMISSION_EMAIL_FLAG:
        permission_msg = None
    elif file_name_wo_ext[-2:].lower() == '-m':
        permission_msg = CORE_MONTHLY_MSG
        # permission_msg = (f"A new MONTHLY ROV-WIDE Sincere summary report has been sent to the CORE GROUP. "
        #                   f"Click to open.")
    else:
        permission_msg = CORE_WEEKLY_MSG
        # permission_msg = (f"A new WEEKLY ROV-WIDE Sincere summary report has been sent to the CORE GROUP. "
        #                   f"Click to open.")

    for email in CORE_EMAIL_LIST:
        # print('  - Adding ROV-wide permission for email: ', email)
        logging.getLogger(name='my_logger').info(f"{inspect.stack()[0][3]}  - Adding ROV-wide permission for email: "
                                                 f"', {email}")
        permission_to_drive_file(drive_service, uploaded_file_id, SEND_PERMISSION_EMAIL_FLAG, email, permission_msg)
