import os
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

SMTP_HOST = os.getenv('SMTP_HOST')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))
SMTP_USER = os.getenv('EMAIL_USER')
SMTP_PASS = os.getenv('EMAIL_PASS')
# Comma-separated list in .env, e.g. EMAIL_RECIPIENTS=foo@example.com,bar@example.com
RECIPIENTS = [addr.strip() for addr in os.getenv('EMAIL_RECIPIENTS', '').split(',') if addr.strip()]


def send_job_results(file_paths: list, date_range: str, stats: str = None):
    missing = []
    if not SMTP_HOST:
        missing.append('SMTP_HOST')
    if not SMTP_USER:
        missing.append('SMTP_USER')
    if not SMTP_PASS:
        missing.append('SMTP_PASS')
    if not RECIPIENTS:
        missing.append('EMAIL_RECIPIENTS')

    if missing:
        print(f"Email not sent: missing configuration for {', '.join(missing)}. Please review your .env file")
        return

    body = "Please see the attached files for the selected scraped jobs.\n"
    if stats:
        body += "\n\n" + stats
    body += "\nThanks,\nCalendar Buddy"

    # Build the email
    msg = EmailMessage()
    msg['Subject'] = f"Job list for {date_range}"
    msg['From'] = SMTP_USER
    msg['To'] = ', '.join(RECIPIENTS)
    msg.set_content(body)

    # Attach each file
    for path in file_paths:
        filename = os.path.basename(path)
        with open(path, 'rb') as f:
            data = f.read()
        maintype, subtype = ('application', 'octet-stream')
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)

    # Send via SMTP
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(SMTP_USER, SMTP_PASS)
        smtp.send_message(msg)
