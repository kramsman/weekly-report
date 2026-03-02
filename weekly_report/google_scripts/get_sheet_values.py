"""Fetch a range from a Google Sheet and validate the header row."""

import inspect
import logging
from typing import Any

from uvbekutils import exit_yes


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
