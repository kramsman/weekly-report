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

from weekly_report.create_report_files import create_report_files
from weekly_report.google_scripts.google_scripts import create_google_services
from weekly_report.google_scripts.google_scripts import upload_files
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
from constants import (
    ROOT_PATH, )

setup_loguru("INFO", "DEBUG")

logger.info(f"({ROOT_PATH=}")

def main():
    """ create Sincere reports and upload them """


    logger.info(f"starting")

    drive_service, sheet_service = create_google_services()
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
        upload_files(drive_service)


    logger.info(f"All done!")

    exit()


if __name__ == '__main__':

    main()
