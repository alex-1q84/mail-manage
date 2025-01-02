import os
import argparse
from dotenv import load_dotenv
from exchangelib import Credentials, Account, DELEGATE

# Load environment variables
load_dotenv()

# Get Exchange server credentials from .env
USERNAME = os.getenv('EXCHANGE_USERNAME')
PASSWORD = os.getenv('EXCHANGE_PASSWORD')
SERVER = os.getenv('EXCHANGE_SERVER')

def tidy_mail(log_file='mail_attachments.log', delete_log_file='deleted_mails.log'):
    """Process mail attachments and log results to files.

    Args:
        log_file (str): Path to file for storing mail details
        delete_log_file (str): Path to file for storing deletion records
    """
    # Set up credentials and account
    credentials = Credentials(username=USERNAME, password=PASSWORD)
    account = Account(
        primary_smtp_address=SERVER,
        credentials=credentials,
        autodiscover=True,
        access_type=DELEGATE
    )

    # Walk through inbox
    for item in account.inbox.all().order_by('-datetime_received'):
        if item.attachments:
            total_size = 0
            attachment_names = []
            has_excel = False

            # Calculate total size and check for Excel files
            for attachment in item.attachments:
                total_size += attachment.size
                attachment_names.append(attachment.name)
                if attachment.name.lower().endswith(('.xls', '.xlsx')):
                    has_excel = True

            # Convert and format size appropriately
            if total_size >= 1024 * 1024:  # >= 1MB
                size_str = f"{total_size / (1024 * 1024):.2f} MB"
            else:  # < 1MB
                size_str = f"{total_size / 1024:.2f} KB"

            # Log mail details
            with open(log_file, 'a') as logger:
                log = (f"Time: {item.datetime_received}\t"
                       + f"Subject: {item.subject}\t"
                       + f"Total Size: {size_str}\t"
                       + f"Attachments: {', '.join(attachment_names)}\n")
                logger.write(log)
                print(log)

            # Delete condition and logging
            if total_size > 1024 * 1024 and has_excel:  # > 1MB
                with open(delete_log_file, 'a') as del_logger:
                    log = (f"Deleted at: {item.datetime_received}\t"
                           + f"Subject: {item.subject}\t"
                           + f"Total Size: {size_str}\n")
                    del_logger.write(log)
                    print(log)
                item.delete()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process mail attachments and delete large Excel files.')
    parser.add_argument('--log', default='mail_attachments.log',
                       help='Path to file for storing mail details')
    parser.add_argument('--delete-log', default='deleted_mails.log',
                       help='Path to file for storing deletion records')

    args = parser.parse_args()

    tidy_mail(log_file=args.log, delete_log_file=args.delete_log)
