""" produce weekly reports on Sincere and upload them to Google drive.
Assign permissions to organizers which sends a Google notification.
 """

# FIXME: Import from the google_scripts directory set up the CFCG directory project.

# TODO speed up the room report upload
# TODO: possible - add chart to all org reports (like admin)
# TODO: Move off my personal google account
# TODO: Add a tab or separate report that shows writers who have never requested.   How to avoid those invite to receive only and not order?
# TODO: add mapping email from organizer to requested email
# TODO: add excluded email list to skip giving permission
# TODO: format all numbers with commas
# TODO: remove pymsgbox references
# TODO when permission can't be granted on google sheet (recipient blocked) log msg and error to log file
# TODO: writer errors assigning permissions to a file rather than pymsgboxes during run

# TODO: add link to error files it's definitely possible with PySide6. Since pyautobek.py uses QLabel, you can use HTML links in the label text combined with the linkActivated signal to
#   open the file in the default editor.
#   The pattern would be:
#   label = QLabel(f'Error saved to: <a href="{filepath}">{filepath}</a>')
#   label.setOpenExternalLinks(False)
#   label.linkActivated.connect(lambda url: subprocess.run(['open', url]))  # macOS
#   On macOS, open filepath.txt opens it in the default app (TextEdit or whatever is set). If you want a specific editor like VS Code: ['code', url].
#   Where to add it: Since pyautobek.py is in your installed .venv, is there a source repo for uvbekutils you maintain? If so, you'd add a new function there —
#   something like alert_with_file_link(msg, filepath, title). If not, you could add the function directly to the installed file or create a wrapper in your project.

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

# FIXME below choked -
if False:  # DOES NOT WORK AS PACKAGE this updates the uvbekutils package which contains the little helper programs
    subprocess.run(["uv", "add", "uvbekutils", "--upgrade-package", "uvbekutils"], check=True)
    # uv add uvbekutils --upgrade-package uvbekutils
    #
    # from terminal
    # uv add uvbekutils@git+https://github.com/kramsman/uvbekutils.git
    # subprocess.run(["uv", "add","uvbekutils@git+https://github.com/kramsman/uvbekutils.git@optiona_showbuttons"],
    #                check=True)
    # subprocess.run(["uv", "add", "uvbekutils", "--upgrade-package", "uvbekutils@optiona_showbuttons"], check=True)

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
    SEND_PERMISSION_EMAIL_FLAG,
    ORG_WEEKLY_MSG,
    CORE_WEEKLY_MSG,
    CORE_MONTHLY_MSG,
    ORG_MONTHLY_MSG,
    OUTPUT_DIR_ADMIN,
    OUTPUT_DIR_REPORTS,
    SINCERE_DOWNLOAD_DIR,
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
        upload_files(drive_service=drive_service, admin_folder_id=ADMIN_REPORT_FOLDER_ID,
                     room_folder_id=ROOM_REPORT_FOLDER_ID, core_email_list=CORE_EMAIL_LIST,
                     test_email_list=TEST_EMAIL_LIST,
                     send_email_flag=SEND_PERMISSION_EMAIL_FLAG, org_weekly_msg=ORG_WEEKLY_MSG,
                     core_weekly_msg=CORE_WEEKLY_MSG, core_monthly_msg=CORE_MONTHLY_MSG,
                     org_monthly_msg=ORG_MONTHLY_MSG, output_dir_admin=OUTPUT_DIR_ADMIN,
                     output_dir_reports=OUTPUT_DIR_REPORTS, sincere_download_dir=SINCERE_DOWNLOAD_DIR)


    logger.info(f"All done!")

    exit()


if __name__ == '__main__':

    main()
