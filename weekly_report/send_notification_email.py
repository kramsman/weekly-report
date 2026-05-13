"""Send Google Drive file notification emails via configurable provider.

Supported providers: "sendgrid", "brevo", "smtp_gmail".
Set EMAIL_PROVIDER in constants.py to switch providers.

NOTE: This module is a candidate to move into uvbekutils once that package is updated,
so it can be shared across projects.
"""

from pathlib import Path

from loguru import logger

from weekly_report.constants import EMAIL_PROVIDER


def send_drive_notification(to_email: str, file_id: str, message: str, subject: str,
                            from_email: str, api_key_file: Path,
                            all_recipients: list[str] | None = None,
                            error_log_file: Path | None = None,
                            error_context: str | None = None) -> None:
    """Send a Drive file notification email via the configured provider.

    Args:
        to_email: Recipient email address.
        file_id: Google Drive file ID — used to construct the link.
        message: Body text to include in the email.
        subject: Email subject line.
        from_email: Sender email address (must be verified with the active provider).
        api_key_file: Path to the plain-text file containing the provider API key or App Password.
        all_recipients: Optional list of all addresses receiving this report (shown in email body).
        error_log_file: Optional path to append errors to instead of raising.
        error_context: Optional context string (e.g. room name) included in error log entries.
    """
    api_key = Path(api_key_file).read_text().strip()
    file_url = f"https://docs.google.com/spreadsheets/d/{file_id}"

    recipients_note = ""
    if all_recipients:
        recipients_note = f"<p><small>This report was sent to: {', '.join(all_recipients)}</small></p>"

    html_content = (
        f"<p>{message.replace(chr(10), '<br>')}</p>"
        f"<p><a href='{file_url}'>Click here to open the report</a></p>"
        f"{recipients_note}"
    )

    try:
        if EMAIL_PROVIDER == "sendgrid":
            _send_sendgrid(from_email, to_email, subject, html_content, api_key)
        elif EMAIL_PROVIDER == "brevo":
            _send_brevo(from_email, to_email, subject, html_content, api_key)
        elif EMAIL_PROVIDER == "smtp_gmail":
            _send_smtp_gmail(from_email, to_email, subject, html_content, api_key)
        else:
            raise ValueError(f"Unknown EMAIL_PROVIDER: {EMAIL_PROVIDER!r}")
        logger.info(f"Email sent ({EMAIL_PROVIDER}): {subject} → {to_email}")
    except Exception as e:
        logger.error(f"Email failed ({EMAIL_PROVIDER}): {subject} → {to_email}: {e}")
        if error_log_file:
            context_str = f"room: {error_context}, " if error_context else ""
            with open(error_log_file, 'a') as f:
                f.write(f"EMAIL ERROR: {context_str}email: {to_email}, subject: {subject}, error: {e}\n")


def _send_sendgrid(from_email: str, to_email: str, subject: str, html_content: str, api_key: str) -> None:
    import sendgrid
    from sendgrid.helpers.mail import Mail
    email = Mail(from_email=from_email, to_emails=to_email, subject=subject, html_content=html_content)
    sg = sendgrid.SendGridAPIClient(api_key=api_key)
    response = sg.send(email)
    logger.debug(f"SendGrid status: {response.status_code}")


def _send_brevo(from_email: str, to_email: str, subject: str, html_content: str, api_key: str) -> None:
    import brevo
    from brevo.transactional_emails.types.send_transac_email_request_to_item import SendTransacEmailRequestToItem
    from brevo.transactional_emails.types.send_transac_email_request_sender import SendTransacEmailRequestSender

    client = brevo.Brevo(api_key=api_key)
    client.transactional_emails.send_transac_email(
        to=[SendTransacEmailRequestToItem(email=to_email)],
        sender=SendTransacEmailRequestSender(email=from_email),
        subject=subject,
        html_content=html_content,
    )


def _send_smtp_gmail(from_email: str, to_email: str, subject: str, html_content: str, api_key: str) -> None:
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = from_email
    msg['To'] = to_email
    msg.attach(MIMEText(html_content, 'html'))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(from_email, api_key)  # api_key is the Gmail App Password
        smtp.sendmail(from_email, to_email, msg.as_string())
