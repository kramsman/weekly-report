"""Interactively upload admin and/or room reports to Google Drive."""

from pathlib import Path
from typing import Any

import pandas as pd
from uvbekutils import pyautobek
from uvbekutils import select_file

from weekly_report.constants import CORE_EMAIL_LIST
from weekly_report.constants import CORE_MONTHLY_MSG
from weekly_report.constants import CORE_WEEKLY_MSG
from weekly_report.constants import ORG_MONTHLY_MSG
from weekly_report.constants import ORG_WEEKLY_MSG
from weekly_report.constants import OUTPUT_DIR_ADMIN
from weekly_report.constants import OUTPUT_DIR_REPORTS
from weekly_report.constants import SINCERE_DOWNLOAD_DIR
from weekly_report.google_scripts.upload_admin_report import upload_admin_report
from weekly_report.google_scripts.upload_room_reports import upload_room_reports


def upload_files(drive_service: Any) -> None:
    """Interactively upload admin and/or room reports to Google Drive.

    Prompts the user to confirm email message text, then offers separate
    choices to upload the admin report and/or a room report directory. Reads
    a Sincere all-users CSV to build the organizer email list for room
    report permissions.

    Args:
        drive_service: Authenticated Google Drive API service resource.
    """

    choice = pyautobek.confirm(f"\nAre these email messages ok to use?\n\n"
                      f"Org Weekly:\n'{ORG_WEEKLY_MSG}'\n\n"
                      f"Admin Weekly:\n'{CORE_WEEKLY_MSG}'\n\n"
                      f"Admin Monthly:\n'{CORE_MONTHLY_MSG}'\n\n"
                      f"Org Monthly:\n'{ORG_MONTHLY_MSG}'\n\n"
                      f"Admin Email Addresses:\n{CORE_EMAIL_LIST}\n\n",
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
                                                          OUTPUT_DIR_ADMIN,
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
                                                    OUTPUT_DIR_REPORTS,
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
                                            SINCERE_DOWNLOAD_DIR,
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
        upload_admin_report(drive_service, str_admin_report_to_upload)

    if upload_room:
        # do the upload from local to google sheets
        upload_room_reports(drive_service, str_report_dir_to_upload, organizer_email_list)
