""" produce weekly reports on Sincere and upload them to Google drive.
Assign permissions to organizers which sends a Google notification.

If you get an error "Gitupdater not found:
  1. go into terminal
  2. copy and enter: source .venv/bin/activate
  3. then enter: uv pip install git+https://github.com/kramsman/gitupdater.git
 """

# FIXME: Import from the google_scripts directory set up the CFCG directory project.

# TODO speed up the room report upload
# TODO: possible - add chart to all org reports (like admin)
# TODO: Move off my personal google account
# TODO: Add a tab or separate report that shows writers who have never requested.   How to avoid those invite to receive only and not order?
# TODO: add mapping email from organizer to requested email
# TODO: add excluded email list to skip giving permission
# TODO: format all numbers with commas

# run gitupdater to make sure bekutils and bekgoogle utility libraries are updated
import sys
import os
sys.path.append(os.path.expanduser("~/Dropbox/Postcard Files/"))
if True:
    import gitupdater

from weekly_report.create_report_files import create_report_files
from bekgoogle import create_google_services_serviceaccount
from weekly_report.upload_files import upload_files
from weekly_report.reminder import reminder

# from apiclient.discovery import build
from loguru import logger

# from apiclient.http import MediaFileUpload  # needed even if Pycharm says not
# import google.auth.transport.requests
from uvbekutils import setup_loguru
from uvbekutils import pyautobek
from weekly_report.constants import (
    ROOT_PATH,
    SCOPES,
    SERVICE_ACCOUNT_FILE,
    ADMIN_REPORT_FOLDER_ID,
    ROOM_REPORT_FOLDER_ID,
    CORE_EMAIL_LIST,
    TEST_EMAIL_LIST,
    TEST_ROOM_LIMIT,
    SEND_PERMISSION_EMAIL_FLAG,
    SENDGRID_API_KEY_FILE,
    SENDGRID_FROM_EMAIL,
    ORG_WEEKLY_MSG,
    ORG_WEEKLY_SUBJECT,
    CORE_WEEKLY_MSG,
    CORE_WEEKLY_SUBJECT,
    CORE_MONTHLY_MSG,
    CORE_MONTHLY_SUBJECT,
    ORG_MONTHLY_MSG,
    ORG_MONTHLY_SUBJECT,
    OUTPUT_DIR_ADMIN,
    OUTPUT_DIR_REPORTS,
    SINCERE_DOWNLOAD_DIR,
    ERROR_LOG_FILE,
)

setup_loguru("INFO", "DEBUG")

logger.info(f"({ROOT_PATH=}")

def main() -> None:
    """Orchestrate Sincere report generation and Google Drive upload.

    Prompts the user to download required Sincere data, then either creates
    local Excel reports (admin-wide and per-room) or uploads previously
    created reports to Google Drive with organizer permission notifications.
    """


    logger.info(f"starting")

    drive_service, sheet_service = create_google_services_serviceaccount(SERVICE_ACCOUNT_FILE, SCOPES)
    logger.info("finished creating google services")

    # Identify which files should be downloaded before starting
    reminder()

    if True:  # False out prompt for testing
        choice = pyautobek.confirm(
                          (f"\n'Run' to create the admin report and room reports locally\n\n"
                               f"'Upload' to copy either the admin report, room reports, or both to Google sheets and send "
                               f"notifications to Organizers\n\n"
                               f"'Exit' to start over"
                               ),
                          'Run, Upload or Exit?',
                          buttons=["Run", 'Upload', 'Exit'])
    else:
        choice = 'run'

    if choice == "exit":
        exit()

    elif choice == "run":
        create_report_files()

    elif choice == "upload":
        if TEST_EMAIL_LIST:
            if pyautobek.confirm(
                f"TEST MODE: Permission emails will only be sent to:\n{TEST_EMAIL_LIST}\n\nContinue?",
                "Test Mode Active",
                ["Continue", "Exit"]) == "exit":
                exit()
        # clear any error log from a previous run
        if ERROR_LOG_FILE.exists():
            ERROR_LOG_FILE.unlink()
        upload_files(drive_service=drive_service, admin_folder_id=ADMIN_REPORT_FOLDER_ID,
                     room_folder_id=ROOM_REPORT_FOLDER_ID, core_email_list=CORE_EMAIL_LIST,
                     test_email_list=TEST_EMAIL_LIST,
                     test_room_limit=TEST_ROOM_LIMIT,
                     send_email_flag=SEND_PERMISSION_EMAIL_FLAG,
                     org_weekly_msg=ORG_WEEKLY_MSG, org_weekly_subject=ORG_WEEKLY_SUBJECT,
                     core_weekly_msg=CORE_WEEKLY_MSG, core_weekly_subject=CORE_WEEKLY_SUBJECT,
                     core_monthly_msg=CORE_MONTHLY_MSG, core_monthly_subject=CORE_MONTHLY_SUBJECT,
                     org_monthly_msg=ORG_MONTHLY_MSG, org_monthly_subject=ORG_MONTHLY_SUBJECT,
                     output_dir_admin=OUTPUT_DIR_ADMIN,
                     output_dir_reports=OUTPUT_DIR_REPORTS, sincere_download_dir=SINCERE_DOWNLOAD_DIR,
                     sendgrid_api_key_file=SENDGRID_API_KEY_FILE, sendgrid_from_email=SENDGRID_FROM_EMAIL)
        if ERROR_LOG_FILE.exists():
            pyautobek.alert_with_file_link("Errors occurred during upload. See details:",
                                           ERROR_LOG_FILE, "Upload Errors")


    logger.info(f"All done!")

    exit()


if __name__ == '__main__':

    main()
