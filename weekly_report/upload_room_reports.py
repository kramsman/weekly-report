"""Upload all room report .xlsx files from a local directory to Google Drive."""

import glob
import os
from pathlib import Path
from typing import Any

import pandas as pd

from bekgoogle.create_drive_subfolder import create_drive_subfolder
from bekgoogle.delete_list_of_google_files import delete_list_of_google_files
from bekgoogle.get_google_file_or_folder_ids import get_google_file_or_folder_ids
from bekgoogle.permission_to_drive_file import permission_to_drive_file
from bekgoogle.upload_sheet_to_drive import upload_sheet_to_drive


def upload_room_reports(drive_service: Any, str_dir_to_upload: str,
                        organizer_email_list: list[list[str]],
                        folder_id: str, send_email_flag: bool,
                        weekly_msg: str | None, monthly_msg: str | None) -> None:
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

    for index, row in report_files_df.iterrows():
        file_w_path, file_name, file_name_wo_ext, room_name = row
        # # assign report_files_df row to variables to avoid report_files_df(loc,'field') notation
        print(f"\nUploading filename:  {file_name}, room_name: {room_name}")

        uploaded_file_id = upload_sheet_to_drive(drive_service, file_w_path, new_folder_id)

        for org, email in organizer_email_list:
            # email = 'briank@kramericore.com'
            if org.replace(",", "-") == room_name:  # had to replace , with - in file names so this makes room match.
                print(f"  - Adding permission for room_name: {room_name}, email: {email}")

                # set permissions and email admin report to Core group
                if not send_email_flag:
                    permission_msg = None
                elif file_name_wo_ext[-2:].lower() == '-m':
                    permission_msg = monthly_msg
                else:
                    permission_msg = weekly_msg

                # TODO: Could excluding emails from upload be done here?

                permission_to_drive_file(drive_service, uploaded_file_id, send_email_flag, email, permission_msg)
                a=1
