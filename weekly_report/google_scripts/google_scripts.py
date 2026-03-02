"""Google Drive and Sheets API helpers for uploading reports and granting permissions."""

import glob
import inspect
import logging
import os
import re
import sys
import traceback
from pathlib import Path
from typing import Any
from pathlib import Path
from pathlib import Path
from pathlib import Path
from pathlib import Path

import google.auth
import pandas as pd
import pymsgbox
from googleapiclient.discovery import build
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from loguru import logger
from loguru import logger
from loguru import logger
from loguru import logger
from loguru import logger
from loguru import logger
from loguru import logger
from loguru import logger
from loguru import logger
from loguru import logger
from loguru import logger
from loguru import logger
from uvbekutils import exit_yes
from uvbekutils import exit_yes
from uvbekutils import exit_yes
from uvbekutils import pyautobek
from uvbekutils import pyautobek
from uvbekutils import pyautobek
from uvbekutils import pyautobek
from uvbekutils import pyautobek
from uvbekutils import pyautobek
from uvbekutils import select_file
from uvbekutils import select_file
from uvbekutils import select_file

from weekly_report.constants import ADMIN_REPORT_FOLDER_ID
from weekly_report.constants import ADMIN_REPORT_FOLDER_ID
from weekly_report.constants import CORE_EMAIL_LIST
from weekly_report.constants import CORE_EMAIL_LIST
from weekly_report.constants import ORG_MONTHLY_MSG
from weekly_report.constants import ORG_MONTHLY_MSG
from weekly_report.constants import ORG_WEEKLY_MSG
from weekly_report.constants import ORG_WEEKLY_MSG
from weekly_report.constants import OUTPUT_DIR_ADMIN
from weekly_report.constants import OUTPUT_DIR_REPORTS
from weekly_report.constants import ROOM_REPORT_FOLDER_ID
from weekly_report.constants import ROOM_REPORT_FOLDER_ID

from weekly_report.constants import ROOT_PATH
from weekly_report.constants import ROOT_PATH
from weekly_report.constants import ROOT_PATH
from weekly_report.constants import ROOT_PATH
from weekly_report.constants import CORE_MONTHLY_MSG
from weekly_report.constants import CORE_MONTHLY_MSG
from weekly_report.constants import CORE_WEEKLY_MSG
from weekly_report.constants import CORE_WEEKLY_MSG
from google.oauth2.credentials import Credentials

from weekly_report.constants import SCOPES
from weekly_report.constants import SEND_PERMISSION_EMAIL_FLAG
from weekly_report.constants import SEND_PERMISSION_EMAIL_FLAG
from weekly_report.constants import SEND_PERMISSION_EMAIL_FLAG
from weekly_report.constants import SEND_PERMISSION_EMAIL_FLAG
from weekly_report.constants import SINCERE_DOWNLOAD_DIR


def get_creds(scopes: list[str], cred_file: str | None = None, cred_dir: Path | None = None, token_file: str | None = None, token_dir: Path | None = None, always_create: bool = False,
              write_token: bool = True) -> Credentials:
    """Obtain valid Google OAuth2 credentials, refreshing or creating as needed.

    Reads an existing token file if available, refreshes an expired access
    token, or runs the full OAuth flow from a credentials JSON file. Alerts
    the user via pymsgbox if credential files are missing.

    Args:
        scopes: OAuth2 scopes to request.
        cred_file: Credentials JSON filename relative to cred_dir.
            Defaults to '../credentials.json'.
        cred_dir: Directory containing the credentials file.
            Defaults to ROOT_PATH.
        token_file: Token JSON filename relative to token_dir.
            Defaults to '../token.json'.
        token_dir: Directory containing the token file.
            Defaults to ROOT_PATH.
        always_create: Reserved for future use; not currently implemented.
            Defaults to False.
        write_token: If True, write the refreshed token back to disk.
            Defaults to True.

    Returns:
        A valid google.oauth2.credentials.Credentials instance.
    """

    # import pathlib
    # import pymsgbox
    # from google.auth.transport.requests import Request
    import google.auth.transport.requests
    from google_auth_oauthlib.flow import InstalledAppFlow

    def bek_cred_flow():
        """Run the installed-app OAuth flow and return new credentials.

        Prompts the user via pymsgbox before opening the browser, then calls
        InstalledAppFlow.run_local_server().

        Returns:
            New credentials returned by the OAuth consent screen flow.
        """
        logger.debug("get_creds using credentials -going to call 'InstalledAppFlow.from_client_secrets_file'")
        flow = InstalledAppFlow.from_client_secrets_file(cred_file, scopes)
        logger.debug(f"got flow: {flow.__dict__=}")
        # creds = flow.run_local_server(port=0)
        msg = ("Calling the 2nd part of flow, 'flow.run_local_server(port=0)'"
                 "\n\n   - Use 'TECH@CENTERFORCOMMONGROUND.ORG' google login."
                 "\n\n   - Must say process completed - close this window."
                 "\n   - If you get error 403, hit back in the browser and try again")
        logger.debug(msg)
        pymsgbox.alert(msg,"Verify Google Account")
        creds = flow.run_local_server(port=0)
        logger.debug(f"{creds.__dict__=}")
        return creds

    logger.debug(f"in get_creds with: {scopes=}, {cred_file=}, {cred_dir=}, {token_file=}, {token_dir=},"
                 f" {always_create=}, {write_token=}")
    if cred_dir is None:
        cred_dir = ROOT_PATH
    if cred_file is None:
        cred_file = '../credentials.json'
    cred_w_path = cred_dir / cred_file

    if token_dir is None:
        token_dir = ROOT_PATH
    if token_file is None:
        token_file = '../token.json'
    token_w_path = token_dir / token_file

    # token_w_path = Path(token_dir) / 'token.json'

    # SCOPES = ['https://www.googleapis.com/auth/drive']
    # TODO: Add code to delete token.json if creds fails.

    # FIXME this should check token (w refresh?) then credential
    if not cred_w_path.is_file():
        pymsgbox.alert(f"'{cred_w_path}' is missing. Copy it from another dir or download from API manager")
        # exit()

    if not token_w_path.is_file() and not cred_w_path.is_file():
        logger.error(f"'{token_w_path}' and {cred_w_path} are both missing.")
        pymsgbox.alert(f"'{token_w_path}' and {cred_w_path} are both missing.")
        exit()

    creds = None

    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if token_w_path.is_file():
        # Credentials.from_authorized_user_file. Creates a Credentials instance from parsed authorized user info.
        # does it work from credentials, token, or either?
        # from_authorized_user_file - Creates a Credentials instance from an authorized user json file.
        creds = Credentials.from_authorized_user_file(token_w_path, scopes)
        logger.error(f"get_creds using token: {creds.__dict__=}")
        # logger.error(f"{dir(creds)=}")  # only names, not values, and includes builtins
        # logger.error(f"{vars(creds)=}")  # same as __dict__
        # logger.error(f"{help(creds)=}")
        # exit()

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                # refresh(self, request) Refreshes the access token.
                # creds.refresh(Request())
                request = google.auth.transport.requests.Request()
                creds.refresh(request)
                logger.error("Creds refresh worked using token")
                logger.debug(f"{creds.__dict__=}")
            except Exception as e:
                pymsgbox.alert(f"Creds(request)) did not work using token.\nGoing to prompt for Google login."
                               f"\n\nError=\n{e}",
                               "Google Credential Issue")
                logger.debug("refresh failed - going to run bek_cred_flow")
                credsX = bek_cred_flow()
                pymsgbox.alert(f"Creds(request)) created.\nRerun program and you should not be prompted for login.",
                               "Google Credential Created")
                # exit()  # TODO: why does this error with SIGDEF?  It used to wok.  Explore.
        else:
            credsX = bek_cred_flow()

        # Save the credentials for the next run
        if write_token:
            logger.debug("writing token")
            with open(token_w_path, 'w') as token:
                token.write(creds.to_json())
    logger.debug(f"ready to leave get_creds: {creds.__dict__=}")
    return creds


def get_sheet_values(sheet_service: Any, sheet_range: str, header_strings: str | list[str]) -> list[list]:
    """Fetch a range from a Google Sheet and validate the header row.

    Retrieves the spreadsheet range, pops the header row, and compares it
    (case-insensitive) to header_strings. Calls exit_yes() if headers do not
    match. Spreadsheet ID is hardcoded in CFCG_LOOKUP_GOOGLE_SHEET_ID.

    Args:
        sheet_service: Authenticated Google Sheets API service resource.
        sheet_range: A1 notation range to retrieve, e.g. 'Sheet1!A:Z'.
        header_strings: Expected headers as a comma-separated string or a
            list of strings.

    Returns:
        Row data (excluding the header row) as a list of lists.
    """
    logging.getLogger(name='my_logger').info(f"Starting {inspect.stack()[0][3]}")

    CFCG_LOOKUP_GOOGLE_SHEET_ID = 9999

    if isinstance(header_strings, str):
        header_strings_list = [field.strip().lower() for field in header_strings.split(",")]
    elif isinstance(header_strings, list):
        header_strings_list = [field.strip().lower() for field in header_strings]
    else:
        exit_yes(f"Headers passed to get_sheet_values is not string or list of headers.  It's "
                 f"{type(header_strings)}.")

    sheet_result = sheet_service.spreadsheets().values().get(spreadsheetId=CFCG_LOOKUP_GOOGLE_SHEET_ID, range=sheet_range).execute()
    sheet_values = sheet_result.get('values', [])
    headings = sheet_values.pop(0)  # read header and compare to ensure file format stays the same
    headings = [field.strip().lower() for field in headings]
    if headings != header_strings_list:
        exit_yes(f"Field names in {sheet_range} sheet header record are not {header_strings_list}; \n\nThey are {headings}")

    return sheet_values


def upload_sheet_to_drive(drive_service: Any, file_to_upload_w_path: str | Path, drive_folder_id: str | list[str]) -> str:
    """Upload a local Excel file to Google Drive as a Google Sheet.

    Accepts a single folder ID string or a list of folder IDs as the
    destination. Alerts the user via pyautobek if the conversion or upload
    fails.

    Args:
        drive_service: Authenticated Google Drive API service resource.
        file_to_upload_w_path: Full path to the local .xlsx file.
        drive_folder_id: Google Drive folder ID or list of IDs for the
            destination.

    Returns:
        The Google Drive file ID of the newly created file.
    """
    # allow one folder id or list to be passed
    if isinstance(drive_folder_id, str):
        drive_folder_id = [drive_folder_id]

    file_name_wo_ext = Path(file_to_upload_w_path).stem
    file_metadata = {'name': file_name_wo_ext,
                     'mimeType': 'application/vnd.google-apps.spreadsheet',
                     "parents": drive_folder_id}

    # converts xlsx to google sheet
    try:
        media = MediaFileUpload(file_to_upload_w_path,
                            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except:
        print(f"**** ERROR CONVERTING XLS FILE TO GOOGLE SHEET: Uploaded file name:  {file_name_wo_ext}, "
              f"path: {file_to_upload_w_path}")
        # pymsgbox.alert(f"**** CONVERTING XLS FILE TO GOOGLE SHEET: Uploaded file name:  {file_to_upload_w_path}",
        #            "CHECK PYTHON CONSOLE FOR ERROR")
        pyautobek.alert(f"\n\n**** CONVERTING XLS FILE TO GOOGLE SHEET: Uploaded file name: "
                     f" {file_to_upload_w_path}","CHECK PYTHON CONSOLE FOR ERROR",
                        )

    try:
        file = drive_service.files().create(body=file_metadata,
                                        supportsAllDrives=True,
                                        media_body=media,
                                        fields='id').execute()
    except:
        # if file upload doesn't work
        print(f"**** ERROR WITH UPLOAD: Uploaded file name: {file_name_wo_ext}, path: "
          f"{file_to_upload_w_path}")
        # pymsgbox.alert(f"**** ERROR WITH UPLOAD: Uploaded file name: {file_name_wo_ext}",
        #            "CHECK PYTHON CONSOLE FOR ERROR")
        pyautobek.alert(f"\n\n**** ERROR WITH UPLOAD: Uploaded file name: {file_name_wo_ext}",
                 "CHECK PYTHON CONSOLE FOR ERROR",
                        )

    uploaded_file_id = file.get('id')

    return uploaded_file_id


def permission_to_drive_file(drive_service: Any, drive_file_id: str, email_flag: bool, email: str, permission_msg: str | None) -> None:
    """Grant reader permission on a Google Drive file to one user.

    Strips gmail '+' alias segments from the email before granting. Falls
    back to sendNotificationEmail=True if the first attempt fails, then
    shows a pyautobek alert if both attempts fail.

    Args:
        drive_service: Authenticated Google Drive API service resource.
        drive_file_id: Google Drive file ID to grant permission on.
        email_flag: Whether to send a Google notification email.
        email: Email address of the user to receive reader access.
        permission_msg: Custom message body for the notification email, or
            None to omit a message.
    """
    # use regex to remove all text between '+' and '@gmail' cause fails in permission grant
    email_clean = re.sub('\+[^>]+@gmail.com', '@gmail.com', email, flags=re.IGNORECASE)
    if email_clean != email:
        print(f"'{email}' changed to '{email_clean}' for permission grant.")
        logging.getLogger(name='my_logger').info(f"In {inspect.stack()[0][3]} - '{email}' changed to '{email_clean}' "
                                                 f"for permission grant.")


    payload = {
        "role": "reader",
        "type": "user",
        "emailAddress": email_clean
    }

    if not email_flag:
        permission_msg = None

    # TODO: need to trap adding permission when notification off because people without google accounts error.
    try:
        # If email is not gmail and doesn't have an associated gmail account, notification must be true, below
        drive_service.permissions().create(fileId=drive_file_id,
                                           supportsAllDrives=True,
                                            sendNotificationEmail=email_flag,
                                            emailMessage=permission_msg,
                                            body=payload).execute()
    except Exception as e:
        print('First Except', sys.exc_info()[2])
        traceback.print_exc()
        # Try again with notification set to true
        try:
            drive_service.permissions().create(fileId=drive_file_id,
                                       sendNotificationEmail=True,
                                       emailMessage=permission_msg,
                                       body=payload).execute()
        except Exception as e:
            print(f"Error giving permission, click ok to continue: Email: {email}, Uploaded file id: {drive_file_id}")
            print(sys.exc_info()[2])
            traceback.print_exc()
            # pymsgbox.alert(
            #     f"*****  ERROR GIVING PERMISSION: Email: {email}, Uploaded file id: {drive_file_id}",
            #     "CHECK PYTHON CONSOLE FOR ERROR")
            pyautobek.alert(f"Error giving permission, ok to ignore: (likely blocked by recipient): Email: "
                         f"{email}, Uploaded file id: {drive_file_id}",
                     "Click OK to Continue",
                            )

    a=1


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


def get_google_file_or_folder_ids(drive_service: Any, file_or_folder: str, folder_name: str, parent: str) -> list[str]:
    """Return Drive IDs of all files or folders matching a name inside a parent.

    Handles pagination automatically. Uses a mimeType comparison to
    distinguish files from folders.

    Args:
        drive_service: Authenticated Google Drive API service resource.
        file_or_folder: 'file' to search for non-folder items; 'folder' to
            search for folders.
        folder_name: Name of the item to search for.
        parent: Google Drive ID of the parent folder to search within.

    Returns:
        List of Google Drive IDs matching the search criteria.
    """
    if file_or_folder.lower() == 'folder':
        file_or_folder_compare = '='
    elif file_or_folder.lower() == 'file':
        file_or_folder_compare = '!='
    else:
        exit_yes(f"Google drive file_or_folder parameter of {file_or_folder} not 'file' or folder'")

    query_string = (f"mimeType {file_or_folder_compare} 'application/vnd.google-apps.folder' and "
                    f"trashed=false and "
                    f"name='{folder_name}' and " 
                    f"parents = '{parent}'"
                    )
    # id_list = []
    page_token = None
    while True:
        response = drive_service.files().list(
            q=query_string,
            spaces='drive',
            fields="nextPageToken, files(id, name, trashed, parents)",
            pageToken=page_token).execute()
        print(f"Found old folders with same name- response: {response}")


        id_list = [file.get('id') for file in response.get('files', [])]

        # for file in response.get('files', []):
        #     print((f"Found folder- name: {file.get('name')}, id: {file.get('id')}, "
        #           f"trashed: {file.get('trashed')}, parents: {file.get('parents')}"))

        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break
    return id_list


def delete_list_of_google_files(drive_service: Any, file_id_list: list[str]) -> None:
    """Move a list of Google Drive items to the trash.

    Args:
        drive_service: Authenticated Google Drive API service resource.
        file_id_list: Google Drive file or folder IDs to trash.
    """
    for folderId in file_id_list:
        drive_service.files().update(fileId=folderId, body={'trashed': True}).execute()
        # trash instead of delete moves to trash; delete gone for good


def create_google_services() -> tuple[Any, Any]:
    """Create and return authenticated Google Drive and Sheets API service objects.

    Calls get_creds() with the project SCOPES and credential paths from
    ROOT_PATH.

    Returns:
        A (drive_service, sheet_service) tuple of authenticated API resources.
    """
    creds = get_creds(SCOPES, cred_file='../credentials.json', cred_dir=ROOT_PATH, token_file='../token.json', token_dir=ROOT_PATH, always_create=False,
                      write_token=True)
    try:
        drive_service = build('drive', 'v3', credentials=creds)
        sheet_service = build('sheets', 'v4', credentials=creds)
    except HttpError as err:
        print(err)

    return drive_service, sheet_service


def create_drive_subfolder(drive_service: Any, folder_name: str, drive_folder_id: str | list[str]) -> str:
    """Create a new subfolder on Google Drive and return its ID.

    Args:
        drive_service: Authenticated Google Drive API service resource.
        folder_name: Name for the new subfolder.
        drive_folder_id: Google Drive ID (or list of IDs) of the parent
            folder.

    Returns:
        The Google Drive ID of the newly created subfolder.
    """
    if isinstance(drive_folder_id, str):
        drive_folder_id = [drive_folder_id]

    file_metadata = {
        'name': folder_name,
        'parents': drive_folder_id,
        'mimeType': 'application/vnd.google-apps.folder'}
    file = drive_service.files().create(body=file_metadata, fields='id').execute()
    new_folder_id = file.get('id')
    # print('Folder ID created: %s' % file.get('id'))

    return new_folder_id


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
    excel_report_files_lst = glob.glob(str(report_dir_to_upload / '[!~]*.xlsx'))
    # xlsx files not starting with ~ (temporary files)

    report_files_df = pd.DataFrame(excel_report_files_lst, columns=['file_w_path'])
    report_files_df['file_name'] = report_files_df['file_w_path'].map(lambda x: os.path.basename(x))
    report_files_df['file_name_wo_ext'] = report_files_df['file_w_path'].map(lambda x: Path(x).stem)
    report_files_df['room_name'] = report_files_df['file_name_wo_ext'].map(lambda x: x.split(' Sincere ')[0])
    # report_files_df['room_name'] = report_files_df['file_name_wo_ext'].map(lambda x: x[0:-34])

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
                    # permission_msg = ("A new MONTHLY VoterLetters summary report is available for your room. "
                    #                   "To access the sheet you will need to be logged in to Google.  Do this by using the Chrome browser or by going to google.com in another browser."
                    #                   "Click to open.")
                else:
                    permission_msg = ORG_WEEKLY_MSG
                    # permission_msg = (f"A new WEEKLY VoterLetters summary report is available for your room. "
                    #                   f"To access the sheet you will need to be logged in to Google.  Do this by using the Chrome browser or by going to google.com in another browser."
                    #                   f"Click to open.")

                # print('  - Adding ROV-wide permission for email: ', email)
                # logging.getLogger(name='my_logger').info(f"{inspect.stack()[0][3]} - Adding room permission for email: "
                #     f"{email}")
                # pymsgbox.confirm("Wait a bit and then try")

                # TODO: Could excluding emails from upload be done here?

                permission_to_drive_file(drive_service, uploaded_file_id, SEND_PERMISSION_EMAIL_FLAG, email, permission_msg)
                a=1


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
