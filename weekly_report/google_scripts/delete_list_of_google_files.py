"""Move a list of Google Drive items to the trash."""

from typing import Any


def delete_list_of_google_files(drive_service: Any, file_id_list: list[str]) -> None:
    """Move a list of Google Drive items to the trash.

    Args:
        drive_service: Authenticated Google Drive API service resource.
        file_id_list: Google Drive file or folder IDs to trash.
    """
    for folderId in file_id_list:
        drive_service.files().update(fileId=folderId, body={'trashed': True}).execute()
        # trash instead of delete moves to trash; delete gone for good
