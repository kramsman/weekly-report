"""Grant reader permission on a Google Drive file to one user."""

import inspect
import logging
import re
import sys
import traceback
from typing import Any

from uvbekutils import pyautobek


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
