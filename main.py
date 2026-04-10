import sys
import os
import win32com.client as win32
from enum import Enum
from coldcallocr import scan

class OutlookFolder(Enum):
    olFolderDeletedItems = 3 # The Deleted Items folder
    olFolderOutbox = 4 # The Outbox folder
    olFolderSentMail = 5 # The Sent Mail folder
    olFolderInbox = 6 # The Inbox folder
    olFolderDrafts = 16 # The Drafts folder
    olFolderJunk = 23 # The Junk E-Mail folder

# Get a reference to Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Get the Inbox folder
inbox = outlook.GetDefaultFolder(OutlookFolder.olFolderInbox.value)

# For subfolders
# inbox = inbox.Folders.Item("")

# Apply the restriction
messages = inbox.Items.Restrict("[ReceivedTime]>='2026-03-13'")

# Check if messages exist
if len(messages) == 0:
    print("There aren't any unread messages in this folder")
    exit()

# Initialize text file
output_file = "OutreachResults.txt"
with open(output_file, "w", encoding="utf-8") as file:
    attachment_path = os.path.join(os.getcwd(), f"images")
    if os.path.exists(attachment_path) == False:
        os.mkdir(attachment_path)
    path = attachment_path
    # Loop over messages
    for message in messages:
        sender = message.SenderName.upper()
        date_sent = message.Senton.date()
        email_body = message.Body
        attachments = message.Attachments
        
        file.write(f"Sender: {sender}\nDate Sent: {date_sent}\n\nEmail Content:\n{email_body}\n\n")
        
        if len(attachments) > 0:
            for attachment in attachments:
                attachment_path = os.path.join(path, attachment.FileName)
                attachment.SaveASFile(attachment_path)
                scanned = scan(attachment_path)
                
                if scanned:
                    file.write("Attachment Content:\n")
                    file.write(scanned + "\n\n")  # Directly write the full text

        file.write("-" * 100 + "\n\n")
        print(f"Processed unread email from {repname}")
