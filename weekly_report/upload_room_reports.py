"""Upload all room report .xlsx files from a local directory to Google Drive."""

import glob
import os
from pathlib import Path
from typing import Any
from loguru import logger

import pandas as pd

from bekgoogle.create_drive_subfolder import create_drive_subfolder
from bekgoogle.delete_list_of_google_files import delete_list_of_google_files
from bekgoogle.get_google_file_or_folder_ids import get_google_file_or_folder_ids
from bekgoogle.permission_to_drive_file import permission_to_drive_file
from bekgoogle.upload_sheet_to_drive import upload_sheet_to_drive
from weekly_report.send_notification_email import send_drive_notification


def upload_room_reports(drive_service: Any, str_dir_to_upload: str,
                        organizer_email_list: list[list[str]],
                        folder_id: str, send_email_flag: bool,
                        weekly_msg: str | None, weekly_subject: str | None,
                        monthly_msg: str | None, monthly_subject: str | None,
                        sendgrid_api_key_file: Path, sendgrid_from_email: str,
                        test_room_limit: int = 0,
                        test_email_list: list[str] | None = None,
                        all_organizer_email_list: list[list[str]] | None = None) -> None:
    """Upload all room report .xlsx files from a local directory to Google Drive.

    Deletes any existing Drive folder with the same name, creates a fresh
    subfolder, uploads each .xlsx as a Google Sheet, and grants reader
    permission to the matching organizer email for each room.

    Args:
        drive_service: Authenticated Google Drive API service resource.
        str_dir_to_upload: Local directory path containing the room .xlsx
            report files.
        organizer_email_list: List of [org_name, email] pairs used to match
            each report file to its organizer's email address.
        folder_id: Google Drive folder ID for the room reports destination.
        send_email_flag: Whether to send Google notification emails.
        weekly_msg: Notification message body for weekly reports.
        monthly_msg: Notification message body for monthly reports.
    """

    report_dir_to_upload = Path(str_dir_to_upload)
    folder_name = report_dir_to_upload.parts[-1]  # the last section of the path # TODO change to better function than parts

    # id list of folders to trash(delete)
    # delete_folder_list1 = get_google_folder_ids(drive_service, folder_name, folder_id)
    delete_folder_list = get_google_file_or_folder_ids(drive_service, 'folder', folder_name,
                                                             folder_id)
    # TODO maybe combine delete folder and file and delete_list_of_google_files
    delete_list_of_google_files(drive_service, delete_folder_list)

    # Create sub folder
    new_folder_id = create_drive_subfolder(drive_service, folder_name, folder_id)

    # get list of filenames in computer directory
    # xlsx files not starting with ~ (temporary files)
    excel_report_files_lst = glob.glob(str(report_dir_to_upload / '[!~]*.xlsx'))

    report_files_df = pd.DataFrame(excel_report_files_lst, columns=['file_w_path'])
    report_files_df['file_name'] = report_files_df['file_w_path'].map(lambda x: os.path.basename(x))
    report_files_df['file_name_wo_ext'] = report_files_df['file_w_path'].map(lambda x: Path(x).stem)
    report_files_df['room_name'] = report_files_df['file_name_wo_ext'].map(lambda x: x.split(' Sincere ')[0])

    report_files_df.sort_values(by=['room_name'], ascending=True, inplace=True)

    rooms_processed = 0
    for index, row in report_files_df.iterrows():
        if test_room_limit and rooms_processed >= test_room_limit:
            logger.info(f"TEST MODE: room limit of {test_room_limit} reached, stopping")
            break
        file_w_path, file_name, file_name_wo_ext, room_name = row
        # # assign report_files_df row to variables to avoid report_files_df(loc,'field') notation
        print(f"\nUploading filename:  {file_name}, room_name: {room_name}")

        if test_room_limit and all_organizer_email_list:
            for org, email in all_organizer_email_list:
                if org.replace(",", "-") == room_name and [org, email] not in organizer_email_list:
                    logger.info(f"TEST MODE: skipping email to {email} ({org})")

        uploaded_file_id = upload_sheet_to_drive(drive_service, file_w_path, new_folder_id)

        # determine notification message based on report type
        if file_name_wo_ext[-2:].lower() == '-m':
            notification_msg = monthly_msg
            notification_subject = monthly_subject
        else:
            notification_msg = weekly_msg
            notification_subject = weekly_subject

        if test_email_list:
            # test mode: send to test addresses regardless of organizer match
            logger.debug(f"TEST MODE: granting permission for room: {room_name}, email(s): {test_email_list}")
            for perm_email in test_email_list:
                permission_to_drive_file(drive_service, uploaded_file_id, False, perm_email, None)
                if send_email_flag and notification_msg:
                    send_drive_notification(perm_email, uploaded_file_id, notification_msg,
                                            f"{notification_subject} — {room_name}",
                                            sendgrid_from_email, sendgrid_api_key_file,
                                            all_recipients=test_email_list)
        else:
            for org, email in organizer_email_list:
                if org.replace(",", "-") == room_name:  # had to replace , with - in file names so this makes room match.
                    print(f"  - Adding permission for room_name: {room_name}, email: {email}")
                    logger.debug(f"Adding organizer permission for org: {org}, email: {email}")
                    permission_to_drive_file(drive_service, uploaded_file_id, False, email, None)
                    if send_email_flag and notification_msg:
                        send_drive_notification(email, uploaded_file_id, notification_msg,
                                                f"{notification_subject} — {room_name}",
                                                sendgrid_from_email, sendgrid_api_key_file,
                                                all_recipients=[email])

        rooms_processed += 1
