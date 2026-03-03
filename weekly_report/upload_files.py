"""Interactively upload admin and/or room reports to Google Drive."""

from pathlib import Path
from typing import Any

import pandas as pd
from uvbekutils import pyautobek
from uvbekutils import select_file

from weekly_report.upload_admin_report import upload_admin_report
from weekly_report.upload_room_reports import upload_room_reports


def upload_files(*, drive_service: Any, admin_folder_id: str, room_folder_id: str, core_email_list: list[str],
                 send_email_flag: bool, org_weekly_msg: str, core_weekly_msg: str, core_monthly_msg: str,
                 org_monthly_msg: str, output_dir_admin: str, output_dir_reports: str, sincere_download_dir: str) -> None:
    """Interactively upload admin and/or room reports to Google Drive.

    Prompts the user to confirm email message text, then offers separate
    choices to upload the admin report and/or a room report directory. Reads
    a Sincere all-users CSV to build the organizer email list for room
    report permissions.

    Args:
        * ():
        drive_service: Authenticated Google Drive API service resource.
        admin_folder_id: Google Drive folder ID for admin reports.
        room_folder_id: Google Drive folder ID for room reports.
        core_email_list: Email addresses for admin report permissions.
        send_email_flag: Whether to send Google notification emails.
        org_weekly_msg: Notification message for organizers (weekly).
        core_weekly_msg: Notification message for core group (weekly).
        core_monthly_msg: Notification message for core group (monthly).
        org_monthly_msg: Notification message for organizers (monthly).
        output_dir_admin: Local directory containing admin report files.
        output_dir_reports: Local directory containing room report directories.
        sincere_download_dir: Local directory containing Sincere CSV exports.
    """

    choice = pyautobek.confirm(f"\nAre these email messages ok to use?\n\n"
                      f"Org Weekly:\n'{org_weekly_msg}'\n\n"
                      f"Admin Weekly:\n'{core_weekly_msg}'\n\n"
                      f"Admin Monthly:\n'{core_monthly_msg}'\n\n"
                      f"Org Monthly:\n'{org_monthly_msg}'\n\n"
                      f"Admin Email Addresses:\n{core_email_list}\n\n",
                      "Check Email Messages",   ['Yes', 'No'])
    if choice == "no":
        exit()


    # choice = pymsgbox.confirm("Upload admin report?", "Upload Admin Reports?", ['Yes', 'No'])
    choice = pyautobek.confirm("\n\nUpload Admin Reports?", "Upload admin report?",  ['Yes', 'No'])
    upload_admin = False
    if choice == "yes":
        upload_admin = True
        if True:  # False for testing w hardcoded admin file
            # pymsgbox.alert("Select ROV Admin report to upload to Core")
            str_admin_report_to_upload = select_file("Select ROV ADMIN report to upload to Core Group",
                                                          output_dir_admin,
                                                     'ROVWide*.xlsx',
                                                     ["Select", "Cancel"],
                                                     'file',
                                                     f"Select ROV ADMIN report to upload to Core Group",
                                                     )
            if str_admin_report_to_upload is None:
                exit()
        else:
            str_admin_report_to_upload = "/Users/Denise/Library/CloudStorage/Dropbox/Postcard Files/VL Admin " \
                                          "Reports/TEST ROVWide Sincere Summary 2023-03-14-W.xlsx"


    # choice = pymsgbox.confirm("Upload room reports?", "Upload Room Reports?", ['Yes', 'No'])
    choice = pyautobek.confirm("\n\nUpload Room Reports?", "Upload room reports?", ['Yes', 'No'])
    upload_room = False
    if choice == "yes":
        upload_room = True
        if True:  # False for testing w report directory hardcoded file
            # pymsgbox.alert("Select report directory to UPLOAD")
            str_report_dir_to_upload = select_file("Select a REPORT DIRECTORY to upload",
                                                    output_dir_reports,
                                                   'all-parent-campaigns-requests*',
                                                   ["Select", "Cancel"],
                                                   "dir",
                                                   "Select the REPORT DIRECTORY",
                                                   )
            if str_report_dir_to_upload is None:
                exit()
        else:
            # str_report_dir_to_upload = "/Users/Denise/Library/CloudStorage/Dropbox/Postcard Files/VL Org Reports/TEST all-parent-campaigns-requests-2023-03-14-W"
            str_report_dir_to_upload = "/Users/Denise/Library/CloudStorage/Dropbox/Postcard Files/VL Org Reports/all-parent-campaigns-requests-2023-05-15 GIDEON-W"

        if True:  # True prompts for users file to read; false uses hardcoded test file
            # pymsgbox.alert("Select Sincere export file for USER LIST (all-users-yyyy-mm-dd.csv)")
            sincere_user_file = select_file("Select Sincere export file for USER LIST (all-users-yyyy-mm-dd.csv)",
                                            sincere_download_dir,
                                            'all-users*.csv',
                                            ["Select", "Cancel"],
                                            "file",
                                            "Select the Sincere report file for all users",
                                            )
            if sincere_user_file is None:
                exit()
            df = pd.read_csv(sincere_user_file, usecols=['name', 'email', 'role', 'is_active', 'organization'])
        else:
            df = pd.read_csv('/Users/Denise/Downloads/all-users-2023-05-15 GIDEON.csv',
                             usecols=['name', 'email', 'role', 'is_active', 'organization'])

        df = df.sort_values(['organization', 'name'], ascending=(True, True))
        organizer_email_list = df.loc[(df['role'].isin(['organizer']) & (df['is_active'])),
                                       ['organization', 'email']].values.tolist()

        # TODO: Could excluding emails from upload be done here?

        # organizer_email_list = [
        #     ['NY-NYC and Long Island', 'kramsman+nycorg@gmail.com'], ['NY-NYC and Long Island', 'bkramer@kramericore.com']
        # ]


    if upload_admin:
        # do the upload from local to google sheets
        upload_admin_report(drive_service=drive_service, admin_report_to_upload=str_admin_report_to_upload,
                            folder_id=admin_folder_id, email_list=core_email_list, send_email_flag=send_email_flag,
                            weekly_msg=core_weekly_msg, monthly_msg=core_monthly_msg)

    if upload_room:
        # do the upload from local to google sheets
        upload_room_reports(drive_service, str_report_dir_to_upload, organizer_email_list,
                            folder_id=room_folder_id, send_email_flag=send_email_flag,
                            weekly_msg=org_weekly_msg, monthly_msg=org_monthly_msg)
