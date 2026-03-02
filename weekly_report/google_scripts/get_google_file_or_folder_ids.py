"""Return Drive IDs of all files or folders matching a name inside a parent."""

from typing import Any

from uvbekutils import exit_yes


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
