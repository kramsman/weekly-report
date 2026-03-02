"""Create and return authenticated Google Drive and Sheets API service objects."""

from typing import Any

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from weekly_report.constants import ROOT_PATH
from weekly_report.constants import SCOPES
from weekly_report.google_scripts.get_creds import get_creds


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
