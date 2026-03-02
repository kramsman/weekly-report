"""Obtain valid Google OAuth2 credentials, refreshing or creating as needed."""

import logging
from pathlib import Path

import pymsgbox
from loguru import logger
from google.oauth2.credentials import Credentials

from weekly_report.constants import ROOT_PATH


def get_creds(scopes: list[str], cred_file: str | None = None, cred_dir: Path | None = None, token_file: str | None = None, token_dir: Path | None = None, always_create: bool = False,
              write_token: bool = True) -> Credentials:
    """Obtain valid Google OAuth2 credentials, refreshing or creating as needed.

    Reads an existing token file if available, refreshes an expired access
    token, or runs the full OAuth flow from a credentials JSON file. Alerts
    the user via pymsgbox if credential files are missing.

    Args:
        scopes: OAuth2 scopes to request.
        cred_file: Credentials JSON filename relative to cred_dir.
            Defaults to '../credentials.json'.
        cred_dir: Directory containing the credentials file.
            Defaults to ROOT_PATH.
        token_file: Token JSON filename relative to token_dir.
            Defaults to '../token.json'.
        token_dir: Directory containing the token file.
            Defaults to ROOT_PATH.
        always_create: Reserved for future use; not currently implemented.
            Defaults to False.
        write_token: If True, write the refreshed token back to disk.
            Defaults to True.

    Returns:
        A valid google.oauth2.credentials.Credentials instance.
    """

    # import pathlib
    # import pymsgbox
    # from google.auth.transport.requests import Request
    import google.auth.transport.requests
    from google_auth_oauthlib.flow import InstalledAppFlow

    def bek_cred_flow():
        """Run the installed-app OAuth flow and return new credentials.

        Prompts the user via pymsgbox before opening the browser, then calls
        InstalledAppFlow.run_local_server().

        Returns:
            New credentials returned by the OAuth consent screen flow.
        """
        logger.debug("get_creds using credentials -going to call 'InstalledAppFlow.from_client_secrets_file'")
        flow = InstalledAppFlow.from_client_secrets_file(cred_file, scopes)
        logger.debug(f"got flow: {flow.__dict__=}")
        # creds = flow.run_local_server(port=0)
        msg = ("Calling the 2nd part of flow, 'flow.run_local_server(port=0)'"
                 "\n\n   - Use 'TECH@CENTERFORCOMMONGROUND.ORG' google login."
                 "\n\n   - Must say process completed - close this window."
                 "\n   - If you get error 403, hit back in the browser and try again")
        logger.debug(msg)
        pymsgbox.alert(msg,"Verify Google Account")
        creds = flow.run_local_server(port=0)
        logger.debug(f"{creds.__dict__=}")
        return creds

    logger.debug(f"in get_creds with: {scopes=}, {cred_file=}, {cred_dir=}, {token_file=}, {token_dir=},"
                 f" {always_create=}, {write_token=}")
    if cred_dir is None:
        cred_dir = ROOT_PATH
    if cred_file is None:
        cred_file = '../credentials.json'
    cred_w_path = cred_dir / cred_file

    if token_dir is None:
        token_dir = ROOT_PATH
    if token_file is None:
        token_file = '../token.json'
    token_w_path = token_dir / token_file

    # token_w_path = Path(token_dir) / 'token.json'

    # SCOPES = ['https://www.googleapis.com/auth/drive']
    # TODO: Add code to delete token.json if creds fails.

    # FIXME this should check token (w refresh?) then credential
    if not cred_w_path.is_file():
        pymsgbox.alert(f"'{cred_w_path}' is missing. Copy it from another dir or download from API manager")
        # exit()

    if not token_w_path.is_file() and not cred_w_path.is_file():
        logger.error(f"'{token_w_path}' and {cred_w_path} are both missing.")
        pymsgbox.alert(f"'{token_w_path}' and {cred_w_path} are both missing.")
        exit()

    creds = None

    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if token_w_path.is_file():
        # Credentials.from_authorized_user_file. Creates a Credentials instance from parsed authorized user info.
        # does it work from credentials, token, or either?
        # from_authorized_user_file - Creates a Credentials instance from an authorized user json file.
        creds = Credentials.from_authorized_user_file(token_w_path, scopes)
        logger.error(f"get_creds using token: {creds.__dict__=}")
        # logger.error(f"{dir(creds)=}")  # only names, not values, and includes builtins
        # logger.error(f"{vars(creds)=}")  # same as __dict__
        # logger.error(f"{help(creds)=}")
        # exit()

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                # refresh(self, request) Refreshes the access token.
                # creds.refresh(Request())
                request = google.auth.transport.requests.Request()
                creds.refresh(request)
                logger.error("Creds refresh worked using token")
                logger.debug(f"{creds.__dict__=}")
            except Exception as e:
                pymsgbox.alert(f"Creds(request)) did not work using token.\nGoing to prompt for Google login."
                               f"\n\nError=\n{e}",
                               "Google Credential Issue")
                logger.debug("refresh failed - going to run bek_cred_flow")
                credsX = bek_cred_flow()
                pymsgbox.alert(f"Creds(request)) created.\nRerun program and you should not be prompted for login.",
                               "Google Credential Created")
                # exit()  # TODO: why does this error with SIGDEF?  It used to wok.  Explore.
        else:
            credsX = bek_cred_flow()

        # Save the credentials for the next run
        if write_token:
            logger.debug("writing token")
            with open(token_w_path, 'w') as token:
                token.write(creds.to_json())
    logger.debug(f"ready to leave get_creds: {creds.__dict__=}")
    return creds
