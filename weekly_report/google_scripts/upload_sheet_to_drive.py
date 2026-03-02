"""Upload a local Excel file to Google Drive as a Google Sheet."""

from pathlib import Path
from typing import Any

from googleapiclient.http import MediaFileUpload
from uvbekutils import pyautobek


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
