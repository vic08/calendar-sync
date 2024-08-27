import objc
import imaplib
import email
from email.header import decode_header
import os
from datetime import datetime, timedelta
import base64
from Foundation import NSDate, NSURL
from EventKit import EKEventStore, EKEvent, EKRecurrenceRule

def init_mail_client():
    # Replace with your credentials
    username = "your_email@example.com"
    password = "your_password"
    imap_server = "imap.example.com"

    # Connect to the IMAP server
    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(username, password)
    mail.select("inbox")  # Select the mailbox you want to check

    return mail

def dispose_mail_client(mail_client):
    mail_client.logout()

def fetch_mail_inbox(mail_client):
    # Get emails from the last 2 days
    date = (datetime.now() - timedelta(2)).strftime("%d-%b-%Y")
    return mail_client.search(None, f'(SINCE "{date}")')

def extract_ics_attachments(mail_client, data):
    FOLDER_NAME = "ics_files"
    # Create a folder to save the .ics files
    os.makedirs(FOLDER_NAME, exist_ok=True)

    # Iterate over the emails
    for num in data[0].split():
        result, msg_data = mail_client.fetch(num, "(RFC822)")
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        # If the email has multiple parts
        if msg.is_multipart():
            for part in msg.walk():
                # Check if it's an attachment
                if part.get_content_disposition() == "attachment":
                    filename = part.get_filename()
                    if filename and filename.endswith(".ics"):
                        # Decode the file and save it
                        filepath = os.path.join(FOLDER_NAME, filename)
                        with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))
        else:
            # If the email is not multipart but has an attachment
            content_type = msg.get_content_type()
            if content_type == "text/calendar":
                filename = f"{msg['Date']}.ics"
                filepath = os.path.join(FOLDER_NAME, filename)
                with open(filepath, "wb") as f:
                    f.write(msg.get_payload(decode=True))

def main():
    mail_client = init_mail_client()
    result, data = fetch_mail_inbox(mail_client)
    extract_ics_attachments(mail_client, data)