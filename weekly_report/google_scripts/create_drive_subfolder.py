"""Create a new subfolder on Google Drive and return its ID."""

from typing import Any


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
