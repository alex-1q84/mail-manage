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


def setup_account():
    """Set up and return the Exchange account."""
    credentials = Credentials(username=USERNAME, password=PASSWORD)
    return Account(
        primary_smtp_address=SERVER,
        credentials=credentials,
        autodiscover=True,
        access_type=DELEGATE
    )


def get_last_processed_time(last_processed_file='last_processed.txt'):
    """Read and return the last processed timestamp from file."""
    last_processed_time = None
    
    if os.path.exists(last_processed_file):
        with open(last_processed_file, 'r') as f:
            content = f.read().strip()
            if content:
                try:
                    return datetime.datetime.fromisoformat(content)
                except ValueError:
                    pass
    return None


def get_all_inbox_folders(account):
    """Get all folders under the inbox including subfolders."""
    folders_to_process = [account.inbox]
    all_folders = []
    while folders_to_process:
        folder = folders_to_process.pop()
        all_folders.append(folder)
        try:
            folders_to_process.extend(folder.children)
        except:
            pass
    return all_folders


def process_item_attachments(item):
    """Process item attachments and return total size, names, and excel flag."""
    total_size = 0
    attachment_names = []
    has_excel = False
    
    for attachment in item.attachments:
        total_size += attachment.size
        attachment_names.append(attachment.name)
        if attachment.name.lower().endswith(('.xls', '.xlsx')):
            has_excel = True
    return total_size, attachment_names, has_excel


def log_mail_details(item, total_size, attachment_names, log_file):
    """Log mail details to the specified log file."""
    write_header = not os.path.exists(log_file)
    
    with open(log_file, 'a') as logger:
        if write_header:
            logger.write("Time,Subject,Total Size,Attachments\n")
        
        time_str = f'"{item.datetime_received}"'
        subject_str = f'"{item.subject}"' if ',' in item.subject else item.subject
        total_size_str = f'"{format_size(total_size)}"'
        attachments_str = f'"{", ".join(attachment_names)}"'
        
        log = f"{time_str},{subject_str},{total_size_str},{attachments_str}\n"
        logger.write(log)
        print(log)


def delete_mail(item, total_size, delete_log_file):
    """delete mail item and log to the specified file."""
    write_del_header = not os.path.exists(delete_log_file)
    
    with open(delete_log_file, 'a') as del_logger:
        if write_del_header:
            del_logger.write("Received at,Deleted at,Subject,Total Size\n")
        
        received_at_str = f'"{item.datetime_received}"'
        deleted_at_str = f'"{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}"'
        subject_str = f'"{item.subject}"' if ',' in item.subject else item.subject
        total_size_str = f'"{format_size(total_size)}"'
        
        log = f"{received_at_str},{deleted_at_str},{subject_str},{total_size_str}\n"
        del_logger.write(log)
        print(log)
    item.move_to_trash()


def update_last_processed_time(latest_timestamp, last_processed_file='last_processed.txt'):
    """Update the last processed timestamp file."""
    if latest_timestamp is not None:
        with open(last_processed_file, 'w') as f:
            f.write(latest_timestamp.isoformat())


def tidy_mail(log_file='mail_attachments.csv', delete_log_file='deleted_mails.csv'):
    """Process mail attachments and log results to files.

    Args:
        log_file (str): Path to file for storing mail details
        delete_log_file (str): Path to file for storing deletion records
    """
    account = setup_account()
    last_processed_time = get_last_processed_time()
    latest_timestamp = last_processed_time
    
    # Process all folders
    folders = get_all_inbox_folders(account)
    for folder in folders:
        # Get items to process
        if last_processed_time:
            items = folder.all().filter(
                datetime_received__gt=last_processed_time).order_by('datetime_received')
        else:
            items = folder.all().order_by('datetime_received')
        
        for item in items:
            if item.attachments:
                # Update latest timestamp
                if latest_timestamp is None or item.datetime_received > latest_timestamp:
                    latest_timestamp = item.datetime_received
                
                # Process item
                total_size, attachment_names, has_excel = process_item_attachments(item)
                log_mail_details(item, total_size, attachment_names, log_file)
                
                # Check deletion condition
                if total_size > 1024 * 1024 and has_excel:  # > 1MB
                    delete_mail(item, total_size, delete_log_file)
    
    update_last_processed_time(latest_timestamp)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process mail attachments and delete large Excel files.')
    parser.add_argument('--log', default='mail_attachments.csv',
                       help='Path to file for storing mail details')
    parser.add_argument('--delete-log', default='deleted_mails.csv',
                       help='Path to file for storing deletion records')
    # add an argument for processing all mails not just new ones.AI!

    args = parser.parse_args()

    tidy_mail(log_file=args.log, delete_log_file=args.delete_log)
