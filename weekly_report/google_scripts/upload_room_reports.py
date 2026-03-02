"""Upload all room report .xlsx files from a local directory to Google Drive."""

import glob
import os
from pathlib import Path
from typing import Any

import pandas as pd

from weekly_report.constants import ORG_MONTHLY_MSG
from weekly_report.constants import ORG_WEEKLY_MSG
from weekly_report.constants import ROOM_REPORT_FOLDER_ID
from weekly_report.constants import SEND_PERMISSION_EMAIL_FLAG
from weekly_report.google_scripts.create_drive_subfolder import create_drive_subfolder
from weekly_report.google_scripts.delete_list_of_google_files import delete_list_of_google_files
from weekly_report.google_scripts.get_google_file_or_folder_ids import get_google_file_or_folder_ids
from weekly_report.google_scripts.permission_to_drive_file import permission_to_drive_file
from weekly_report.google_scripts.upload_sheet_to_drive import upload_sheet_to_drive


def upload_room_reports(drive_service: Any, str_dir_to_upload: str, organizer_email_list: list[list[str]]) -> None:
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
    """

    report_dir_to_upload = Path(str_dir_to_upload)
    folder_name = report_dir_to_upload.parts[-1]  # the last section of the path # TODO change to better function than parts

    # id list of folders to trash(delete)
    # delete_folder_list1 = get_google_folder_ids(drive_service, folder_name, ROOM_REPORT_FOLDER_ID)
    delete_folder_list = get_google_file_or_folder_ids(drive_service, 'folder', folder_name,
                                                             ROOM_REPORT_FOLDER_ID)
    # TODO maybe combine delete folder and file and delete_list_of_google_files
    delete_list_of_google_files(drive_service, delete_folder_list)

    # Create sub folder
    new_folder_id = create_drive_subfolder(drive_service, folder_name, ROOM_REPORT_FOLDER_ID)

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
                if not SEND_PERMISSION_EMAIL_FLAG:
                    permission_msg = None
                elif file_name_wo_ext[-2:].lower() == '-m':
                    permission_msg = ORG_MONTHLY_MSG
                else:
                    permission_msg = ORG_WEEKLY_MSG

                # TODO: Could excluding emails from upload be done here?

                permission_to_drive_file(drive_service, uploaded_file_id, SEND_PERMISSION_EMAIL_FLAG, email, permission_msg)
                a=1
