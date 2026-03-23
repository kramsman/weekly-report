"""Send Google Drive file notification emails via SendGrid.

NOTE: This module is a candidate to move into uvbekutils once that package is updated,
so it can be shared across projects.
"""

from pathlib import Path

import sendgrid
from sendgrid.helpers.mail import Mail
from loguru import logger


def send_drive_notification(to_email: str, file_id: str, message: str, subject: str,
                            from_email: str, api_key_file: Path,
                            all_recipients: list[str] | None = None) -> None:
    """Send a Drive file notification email via SendGrid.

    Args:
        to_email: Recipient email address.
        file_id: Google Drive file ID — used to construct the link.
        message: Body text to include in the email.
        subject: Email subject line.
        from_email: Sender email address (must be verified in SendGrid).
        api_key_file: Path to the plain-text file containing the SendGrid API key.
        all_recipients: Optional list of all addresses receiving this report (shown in email body).
    """
    api_key = Path(api_key_file).read_text().strip()
    file_url = f"https://docs.google.com/spreadsheets/d/{file_id}"

    recipients_note = ""
    if all_recipients:
        recipients_note = f"<p><small>This report was sent to: {', '.join(all_recipients)}</small></p>"

    email = Mail(
        from_email=from_email,
        to_emails=to_email,
        subject=subject,
        html_content=f"<p>{message.replace('\n', '<br>')}</p><p><a href='{file_url}'>Click here to open the report</a></p>{recipients_note}",
    )

    sg = sendgrid.SendGridAPIClient(api_key=api_key)
    try:
        response = sg.send(email)
        logger.info(f"Email sent: {subject} → {to_email} (status {response.status_code})")
    except Exception as e:
        logger.error(f"SendGrid failed: {subject} → {to_email}: {e}")
