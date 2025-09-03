import os
import argparse
from dotenv import load_dotenv
from exchangelib import Credentials, Account, DELEGATE
import datetime

# Load environment variables
load_dotenv()

# Get Exchange server credentials from .env
USERNAME = os.getenv('EXCHANGE_USERNAME')
PASSWORD = os.getenv('EXCHANGE_PASSWORD')
SERVER = os.getenv('EXCHANGE_SERVER')


def format_size(total_size):
    """Convert bytes to human-readable format (KB or MB)."""
    if total_size >= 1024 * 1024:  # >= 1MB
        return f"{total_size / (1024 * 1024):.2f} MB"
    else:  # < 1MB
        return f"{total_size / 1024:.2f} KB"


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

            # Check if log file exists to write headers
            write_header = not os.path.exists(log_file)
            
            with open(log_file, 'a') as logger:
                # Write header if file is new
                if write_header:
                    logger.write("Time,Subject,Total Size,Attachments\n")
                
                # Properly format each field for CSV
                time_str = f'"{item.datetime_received}"'
                subject_str = f'"{item.subject}"' if ',' in item.subject else item.subject
                total_size_str = f'"{format_size(total_size)}"'
                attachments_str = f'"{", ".join(attachment_names)}"'
                
                log = f"{time_str},{subject_str},{total_size_str},{attachments_str}\n"
                logger.write(log)
                print(log)

            # Delete condition and logging
            if total_size > 1024 * 1024 and has_excel:  # > 1MB
                # Check if delete log file exists to write headers
                write_del_header = not os.path.exists(delete_log_file)
                
                with open(delete_log_file, 'a') as del_logger:
                    # Write header if file is new
                    if write_del_header:
                        del_logger.write("Received at,Deleted at,Subject,Total Size\n")
                    
                    # Properly format each field for CSV
                    received_at_str = f'"{item.datetime_received}"'
                    deleted_at_str = f'"{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}"'
                    subject_str = f'"{item.subject}"' if ',' in item.subject else item.subject
                    total_size_str = f'"{format_size(total_size)}"'
                    
                    log = f"{received_at_str},{deleted_at_str},{subject_str},{total_size_str}\n"
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
