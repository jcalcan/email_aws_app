import os
import csv
import boto3
from datetime import datetime
from botocore.exceptions import NoCredentialsError, ClientError
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import tkinter as tk
from tkinter import filedialog

# AWS SES configuration
SENDER = os.environ.get("SENDER")
AWS_REGION = 'us-east-1'
AWS_ACCESS_KEY_ID = os.environ.get("AWS_ACCESS_KEY_ID")
AWS_SECRET_ACCESS_KEY = os.environ.get("AWS_SECRET_ACCESS_KEY")
DEFAULT_PATH = os.environ.get("DEFAULT_PATH")
MAX_EMAILS = 600  # Set the maximum number of emails to send


def load_email_template(file_path):
    with open(file_path, 'r') as file:
        content = file.read().split('\n', 2)
        subject = content[0].replace('SUBJECT:', '').strip()
        body = content[2].strip()
        return subject, body


def send_email(client, to_recipient, bcc_recipient, first_name, property_address, email_count):

    subject_template, body_template = load_email_template('email_template.txt')

    subject = subject_template.format(property_address=property_address)
    body = body_template.format(first_name=first_name)

    message = MIMEMultipart()
    message['Subject'] = subject
    message['From'] = SENDER
    message['To'] = to_recipient

    # Add body to email
    message.attach(MIMEText(body, 'plain'))

    destinations = [to_recipient]
    if bcc_recipient:
        destinations.append(bcc_recipient)

    try:
        response = client.send_raw_email(
            Source=SENDER,
            Destinations=destinations,
            RawMessage={'Data': message.as_string()}
        )
        print(f"Email {email_count} sent to {to_recipient} (To) and {bcc_recipient or 'No BCC'} (Bcc). "
              f"Message ID: {response['MessageId']}")
        return True
    except ClientError as e:
        print(f"Error sending email: {e.response['Error']['Message']}")
        return False


def process_csv(csv_file):
    try:
        # Initialize AWS SES client
        client = boto3.client('ses',
                              region_name=AWS_REGION,
                              aws_access_key_id=AWS_ACCESS_KEY_ID,
                              aws_secret_access_key=AWS_SECRET_ACCESS_KEY)

        with open(csv_file, 'r') as file:
            csv_reader = csv.DictReader(file)
            rows = list(csv_reader)

        email_count = 1
        for row in rows:
            if email_count >= MAX_EMAILS:
                print(f"Maximum email limit ({MAX_EMAILS}) reached. Stopping process.")
                break

            fullname = row.get("Contact1Name", "").split(' ', 1)
            first_name = fullname[0] if fullname else ""
            last_name = fullname[1] if len(fullname) > 1 else ""

            property_address = row.get("PropertyAddress", "")
            email1 = row.get("Contact1Email_1", "")
            email2 = row.get("Contact1Email_2", "")

            if email1 and email2:
                to_recipient = email1
                bcc_recipient = email2
            elif email1:
                to_recipient = email1
                bcc_recipient = None
            elif email2:
                to_recipient = email2
                bcc_recipient = None
            else:
                continue

            if first_name and (to_recipient or bcc_recipient):  # Only proceed if there's a valid name and email
                if send_email(client, to_recipient, bcc_recipient, first_name, property_address, email_count):
                    email_count += 1

            else:
                print(f"Skipping email for row with empty name: {row}")

        print(f"Email sending process completed. Total emails sent: {email_count}")

    except NoCredentialsError:
        print("AWS credentials not available or invalid.")
    except Exception as e:
        print(f"An error occurred: {e}")


def select_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(
        title="Select CSV file",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        initialdir=DEFAULT_PATH
    )
    return file_path


# Main execution
csv_file = select_file()
while not csv_file:
    print("No file selected. Please choose a CSV file.")
    csv_file = select_file()

print(f"Processing file: {csv_file}")
process_csv(csv_file)
